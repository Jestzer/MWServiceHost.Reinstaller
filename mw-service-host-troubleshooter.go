package main

import (
	"bufio"
	"fmt"
	"os"
	"os/exec"
	"os/user"
	"path/filepath"
	"strings"
	"syscall"

	"github.com/fatih/color"
	"github.com/shirou/gopsutil/process"
)

func main() {

	// Create a color red object.
	red := color.New(color.FgRed)
	redbg := color.New(color.BgRed)
	cyanfg := color.New(color.FgCyan)

	// Add some code for debug mode.

	var processDetected bool
	processDetected = false

	// Let them know we're getting started.
	println("Searching for the MW Service Host process.")

	processName := "MathWorksServiceHost.exe"

	// Get a list of all running processes.
	processes, err := process.Processes()
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	// Loop through to look the list of processes and check if MWSH is running.
	for _, p := range processes {
		name, err := p.Name()
		if err == nil && strings.EqualFold(name, processName) {
			processDetected = true
			fmt.Printf("Process '%s' is running with PID %d\n", name, p.Pid)

			// Prompt the user to end the process.
			fmt.Print("Do you want to end this process? (y/n)\n>> ")
			reader := bufio.NewReader(os.Stdin)
			input, _ := reader.ReadString('\n')
			input = strings.TrimSpace(input)

			if strings.EqualFold(input, "y") {
				err := p.Terminate()
				if err == nil {
					fmt.Printf("%s (PID %d) has been terminated.\n Please wait for the next step.", name, p.Pid)
				} else {
					fmt.Printf("Error terminating the process: %v\n", err)
				}
			} else {
				fmt.Println("Process was not terminated. Exiting the program until you're ready.")
				os.Exit(0)
			}
			break
		}
	}

	if !processDetected {
		fmt.Printf("%s is not running. Skipping termination. \n", processName)
	}

	// Get the Windows username from whoever launched the program.
	currentUser, err := user.Current()
	if err != nil {
		fmt.Printf("Error detecting the Windows username: %v\n", err)
		os.Exit(1)
	}
	username := extractUsername(currentUser.Username)

	// Look for any installed ServiceHosts by searching for "v" folders.
	searchDirectory := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks\\ServiceHost", username)

	var activeFolders []string
	var inactiveFolders []string

	dirs, err := os.ReadDir(searchDirectory)
	if err != nil {
		fmt.Printf("Error reading the directory: %v\n", err)
		os.Exit(1)
	}

	// Find actual installations of the Service Host.
	//fmt.Println("Folders starting with 'v.1' or 'v20':")
	// Add some code to make this check more thorough.
	for _, dir := range dirs {
		if dir.IsDir() {
			dirName := dir.Name()
			if strings.HasPrefix(dirName, "v.1") || strings.HasPrefix(dirName, "v20") {
				dirPath := filepath.Join(searchDirectory, dirName)
				size, err := getFolderSize(dirPath)
				if err != nil {
					fmt.Printf("Error calculating folder size for %s: %v\n", dirName, err)
				} else {
					//fmt.Printf("Folder '%s' has size %d bytes\n", dirName, size)
					if size > 10*1024*1024 { // 10 MB
						//"Active" folders.
						redbg.Printf("Folder '%s' is marked for upcoming deletion.\n", dirName)
						activeFolders = append(activeFolders, dirPath)
					} else {
						// "Inactive" folders.
						redbg.Printf("Folder '%s' is marked for upcoming deletion.\n", dirName)
						inactiveFolders = append(inactiveFolders, dirPath)
					}
				}
			}
		}
	}

	// Run UninstallMathWorksServiceHost.exe for each "active" folder.
	if len(activeFolders) > 0 || len(inactiveFolders) > 0 {
		red.Print("Are you ready to uninstall any installed Service Hosts? This action CANNOT BE REVERSED! ")
		red.Print("Type \"uninstall\" to confirm. Type any anything else to cancel.\n>> ")
		uninstallReader := bufio.NewReader(os.Stdin)
		uninstallInput, _ := uninstallReader.ReadString('\n')
		uninstallInput = strings.TrimSpace(uninstallInput)
		if strings.EqualFold(uninstallInput, "uninstall") {
			fmt.Println("Uninstallation in progress.")
			for _, folder := range activeFolders {
				executablePath := filepath.Join(folder, "mci", "bin", "win64", "UninstallMathWorksServiceHost.exe")
				cmd := exec.Command(executablePath)
				cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true} // Hide the command window
				err := cmd.Start()
				if err != nil {
					fmt.Printf("Error starting UninstallMathWorksServiceHost.exe: %v\n", err)
					fmt.Print(" Exiting the program.")
					os.Exit(1)
				} else {
					err := cmd.Wait()
					if err != nil {
						fmt.Printf("Error waiting for UninstallMathWorksServiceHost.exe to finish: %v\n", err)
					}
					fmt.Printf("UninstallMathWorksServiceHost.exe has finished running for %s.\n", folder)
				}
			}
		} else {
			fmt.Println("Service Host not uninstalled. Exiting the program.")
			os.Exit(1)
		}
	} else {
		cyanfg.Println("No installed Service Hosts detected! Exiting the program.")
		os.Exit(2)
	}
	// Ask the user if they are ready to delete the old ServiceHost folders
	red.Print("Are you ready to delete the old ServiceHost folder, MATLAB Connector folder, their contents, and all other related files? ")
	red.Print("Any related content and folders marked for deletion above will be PERMANENTLY deleted! Type \"delete\" to confirm. ")
	red.Print("Type any anything else to cancel.\n>> ")
	reader := bufio.NewReader(os.Stdin)
	input, _ := reader.ReadString('\n')
	input = strings.TrimSpace(input)

	// Delete folders starting with "v.1", "v2"' and "LatestInstall.Info" file.
	if strings.EqualFold(input, "delete") {
		fmt.Println("Deletion in progress. Please wait.")
		for _, folder := range activeFolders {
			err := os.RemoveAll(folder)
			if err != nil {
				fmt.Printf("Error deleting folder %s: %v\n", folder, err)
			}
		}

		// Delete the inactive folders as well.
		for _, folder := range inactiveFolders {
			err := os.RemoveAll(folder)
			if err != nil {
				fmt.Printf("Error deleting folder %s: %v\n", folder, err)
			}
		}

		// Delete LatestInstall.Info in the searchDirectory.
		latestInstallInfoPath := filepath.Join(searchDirectory, "LatestInstall.Info")
		err := os.Remove(latestInstallInfoPath)
		if err != nil {
			fmt.Printf("Error deleting LatestInstall.Info: %v\n", err)
		} else {
			fmt.Println("Old ServiceHost folders and LatestInstall.Info deleted.")
		}
	} else {
		fmt.Println("Folders and files were not deleted. Exiting the program.")
		os.Exit(1)
	}

	// Add code to prompt and delete the MATLAB Connector folder.
	// Add code to prompt the user to download the Service Host again.
}

// Function to get rid of your computer's name. We don't care about it.
func extractUsername(fullUsername string) string {
	parts := strings.SplitN(fullUsername, "\\", 2)
	if len(parts) == 2 {
		return parts[1]
	}
	return fullUsername
}

// Function to detect folder size.
func getFolderSize(folderPath string) (int64, error) {
	var folderSize int64

	err := filepath.Walk(folderPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() {
			folderSize += info.Size()
		}
		return nil
	})

	if err != nil {
		return 0, err
	}

	return folderSize, nil
}
