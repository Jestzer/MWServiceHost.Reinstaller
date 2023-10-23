package main

import (
	"bufio"
	"fmt"
	"io"
	"net/http"
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

	// Figure out the folder/directory this program is currently running out of.
	thisProgramDir, errTPD := os.Getwd()
	if errTPD != nil {
		fmt.Printf("Error getting the current working directory: %v\n", errTPD)
		fmt.Println("Exiting the program.")
		os.Exit(1)
	}

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
					fmt.Printf("%s (PID %d) has been terminated.\nPlease wait for the next step.", name, p.Pid)
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
	serviceHostDir := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks\\ServiceHost", username)

	// Define the Local MathWorks folder/directory.
	localMathworksDir := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks", username)

	// Define the new Service Host download directory.
	newServiceHostDir := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\Temp", username)

	var activeFolders []string
	var inactiveFolders []string

	dirs, err := os.ReadDir(serviceHostDir)
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
				dirPath := filepath.Join(serviceHostDir, dirName)
				size, err := getFolderSize(dirPath)
				if err != nil {
					fmt.Printf("Error calculating folder size for %s: %v\n", dirName, err)
				} else {
					//fmt.Printf("Folder '%s' has size %d bytes\n", dirName, size)
					if size > 10*1024*1024 { // 10 MB
						//"Active" folders.
						redbg.Printf("\nThe folder '%s' is marked for being moved.", dirName)
						activeFolders = append(activeFolders, dirPath)
					} else {
						// "Inactive" folders.
						redbg.Printf("\nThe folder '%s' is marked for being moved.", dirName)
						inactiveFolders = append(inactiveFolders, dirPath)
					}
				}
			}
		}
	}

	// Run UninstallMathWorksServiceHost.exe for each "active" folder.
	if len(activeFolders) > 0 || len(inactiveFolders) > 0 {
		red.Print("\nAre you ready to uninstall any installed Service Hosts? This action CANNOT BE REVERSED! ")
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
		// Add some code to detect if the MATLABConnector folder is here.
		cyanfg.Println("No installed Service Hosts detected! Exiting the program.")
		os.Exit(2)
	}
	// Ask the user if they are ready to move the old ServiceHost folders
	fmt.Println("\nAre you ready to move the old ServiceHost folder, MATLAB Connector folder, their contents, and all other related files?")
	red.Print("Any related content and folders marked for being moved above will be moved into a folder called \"Old_service_host_files\"")
	red.Print(". You are responsible for deleting this folder (if you wish) after the old Service Host files are moved into this folder.")
	red.Print("\nType \"move\" to confirm. ")
	red.Print("Type any anything else to cancel.\n>> ")
	reader := bufio.NewReader(os.Stdin)
	input, _ := reader.ReadString('\n')
	input = strings.TrimSpace(input)

	// Move folders starting with "v.1", "v2"' and "LatestInstall.Info" file.
	if strings.EqualFold(input, "move") {
		fmt.Println("Moving in progress. Please wait.")

		// Move "active" folders into the old Service Host files folder.
		if len(activeFolders) > 0 {
			if err := moveItemsToOldServiceHostFilesDir(activeFolders); err != nil {
				fmt.Printf("Error moving active folders to backup: %v\n", err)
				os.Exit(1)
			}
		}

		// Move "inactive" folders into the old Service Host files folder.
		if len(inactiveFolders) > 0 {
			if err := moveItemsToOldServiceHostFilesDir(inactiveFolders); err != nil {
				fmt.Printf("Error moving inactive folders to backup: %v\n", err)
				os.Exit(1)
			}
		}

		// Move LatestInstall.Info into the old Service Host files folder.
		latestInstallInfoPath := filepath.Join(serviceHostDir, "LatestInstall.Info")
		if _, err := os.Stat(latestInstallInfoPath); err == nil {
			dest := filepath.Join(thisProgramDir, "Old_service_host_files", "LatestInstall.Info")
			err := os.Rename(latestInstallInfoPath, dest)
			if err != nil {
				fmt.Printf("Error moving LatestInstall.Info to backup: %v\n", err)
				os.Exit(1)
			}
		}

		// Move the MATLABConnector folder into the old Service Host files folder.
		// Define the needed paths now.
		oldServiceHostFilesDirName := "Old_service_host_files"
		oldServiceHostFilesDir := filepath.Join(thisProgramDir, oldServiceHostFilesDirName)
		matlabConnectorDir := filepath.Join(localMathworksDir, "MATLABConnector")
		if _, err := os.Stat(matlabConnectorDir); err == nil {
			// Move the MATLABConnector folder to the oldServiceHostFiles directory.
			dest := filepath.Join(oldServiceHostFilesDir, "MATLABConnector")

			// Rename, store, and display the error, if needed.
			if moveConnectorError := os.Rename(matlabConnectorDir, dest); moveConnectorError != nil {
				fmt.Println("Error:", moveConnectorError)
				os.Exit(1)
			}
		}

		fmt.Println("Service Host files and folders moved to 'Old_service_host_files'.")
	} else {
		fmt.Println("Folders and files were not moved. Exiting the program.")
		os.Exit(1)
	}

	// Add code to prompt the user to download the Service Host again.

	fmt.Print("Would you like to download and install the latest Service Host? (y/n): ")
	_, err = fmt.Scan(&input)
	if err != nil {
		fmt.Println("Error reading input:", err)
		return
	}

	if input == "y" || input == "Y" {
		fmt.Println("Downloading latest Senvice Host installer. Please wait.")
		// Define the URL of the installer
		installerURL := "https://ssd.mathworks.com/supportfiles/downloads/MathWorksServiceHost/v2023.10.0.4/installers/mathworksservicehost_2023.10.0.4_win64_installer.exe"

		installerPath := filepath.Join(newServiceHostDir, "mathworksservicehost_installer.exe")

		// Create the destination directory if it doesn't exist
		if err := os.MkdirAll(newServiceHostDir, 0755); err != nil {
			fmt.Println("Error creating directory:", err)
			return
		}

		// Download the installer and save it.
		resp, err := http.Get(installerURL)
		if err != nil {
			fmt.Println("Error downloading installer:", err)
			return
		}
		defer resp.Body.Close()

		file, err := os.Create(installerPath)
		if err != nil {
			fmt.Println("Error creating installer file:", err)
			return
		}
		defer file.Close()

		_, err = io.Copy(file, resp.Body)
		if err != nil {
			fmt.Println("Error saving installer:", err)
			return
		}

		// Run the installer
		cmd := exec.Command(installerPath)
		cmd.Stdout = os.Stdout
		cmd.Stderr = os.Stderr

		err = cmd.Run()
		if err != nil {
			// Add some code to extract the installer, rather than just running it, since that doesn't seem to work.
			fmt.Println("Error running the installer:", err)
		} else {
			fmt.Println("Installation completed.")
		}
	} else if input == "n" || input == "N" {
		fmt.Println("Installation canceled.")
	} else {
		fmt.Println("Invalid input. Please enter 'y' for yes or 'n' for no.")
	}
}

// Function to move old Service Host files into a different folder.
func moveItemsToOldServiceHostFilesDir(items []string) error {
	oldServiceHostFilesDirName := "Old_service_host_files"
	thisProgramDir, err := os.Getwd()
	if err != nil {
		return err
	}

	oldServiceHostFilesDir := filepath.Join(thisProgramDir, oldServiceHostFilesDirName)

	// Create the oldServiceHostFiles directory if it doesn't exist.
	if _, err := os.Stat(oldServiceHostFilesDir); os.IsNotExist(err) {
		if err := os.Mkdir(oldServiceHostFilesDir, os.ModePerm); err != nil {
			return err
		}
	}

	// Move items to the oldServiceHostFiles directory.
	for _, item := range items {
		dest := filepath.Join(oldServiceHostFilesDir, filepath.Base(item))
		err := os.Rename(item, dest)
		if err != nil {
			return err
		}
	}

	return nil
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
