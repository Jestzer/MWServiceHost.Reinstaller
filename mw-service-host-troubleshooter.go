package main

import (
	"bufio"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"os/signal"
	"os/user"
	"path/filepath"
	"strings"
	"syscall"

	"github.com/chzyer/readline"
	"github.com/fatih/color"
	"github.com/shirou/gopsutil/process"
)

func main() {

	redText := color.New(color.FgRed).SprintFunc()
	cyanText := color.New(color.FgCyan).SprintFunc()
	redBackground := color.New(color.BgRed).SprintFunc()

	// Use readline to make text input not suck so much.
	rl, err := readline.New("> ")
	if err != nil {
		panic(err)
	}
	defer rl.Close()

	// Setup for better Ctrl+C messaging. This is a channel to receive OS signals.
	signalChan := make(chan os.Signal, 1)
	signal.Notify(signalChan, os.Interrupt, syscall.SIGTERM)

	// Start a Goroutine to listen for signals.
	go func() {

		// Wait for the signal.
		<-signalChan

		// Handle the signal by exiting the program and reporting it as so.
		fmt.Print(redBackground("\nExiting from user input..."))
		os.Exit(0)
	}()

	// Figure out the folder/directory this program is currently running out of.
	thisProgramDir, errTPD := os.Getwd()
	if errTPD != nil {
		fmt.Printf("Error getting the current working directory: %v\n", errTPD)
		fmt.Println("Exiting the program.")
		os.Exit(1)
	}

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
	serviceHostFolder := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks\\ServiceHost", username)

	// Define the Local MathWorks folder/directory.
	matlabConnectorFolder := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks\\MATLABConnector", username)

	// Define the new Service Host download directory.
	newServiceHostDir := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\Temp", username)

	movedServiceHostFilesFolder := filepath.Join(thisProgramDir, "Old_service_host_files")

	// Ask the user if they are ready to move the old ServiceHost folders
	fmt.Println("\nAre you ready to move the old ServiceHost folder, MATLAB Connector folder, their contents, and all other related files?")
	fmt.Print(cyanText("Any related content and folders marked for being moved above will be moved into a folder called \"Old_service_host_files\""))
	fmt.Print(cyanText(". You are responsible for deleting this folder (if you wish) after the old Service Host files are moved into this folder."))
	fmt.Print(cyanText("\nType \"move\" to confirm. This is case-sensitive."))
	fmt.Print(cyanText("Type any anything else to cancel.\n"))

	for {
		input, err := rl.Readline()
		if err != nil {
			if err.Error() == "Interrupt" {
				fmt.Println(redText("Exiting from user input."))
			} else {
				fmt.Print(redText("Error reading line:", err))
				continue
			}
			return
		}
		input = strings.TrimSpace(input)

		if input == "move" {
			break
		} else {
			fmt.Println(redText("Exiting since the user has declined moving necessary folders."))
			os.Exit(0)
		}
	}

	// Move the ServiceHost folder.
	err = moveFolder(serviceHostFolder, filepath.Join(movedServiceHostFilesFolder, "ServiceHost"))
	if err != nil {
		fmt.Printf(redText("Error moving ServiceHost folder: %s\n", err))
		os.Exit(1)
	} else {
		fmt.Println("ServiceHost folder moved successfully.")
	}

	// Move the MATLABConnector folder.
	err = moveFolder(matlabConnectorFolder, filepath.Join(movedServiceHostFilesFolder, "MATLABConnector"))
	if err != nil {
		fmt.Printf(redText("Error moving MATLABConnector folder: %s\n", err))
		os.Exit(1)
	} else {
		fmt.Println("MATLABConnector folder moved successfully.")
	}

	fmt.Println("Downloading latest Service Host installer. Please wait.")

	installerURL := "https://ssd.mathworks.com/supportfiles/downloads/MathWorksServiceHost/v2024.2.0.3/installers/mathworksservicehost_2024.2.0.3_win64_installer.exe"
	installerPath := filepath.Join(newServiceHostDir, "mathworksservicehost_installer.exe")

	// Create the destination directory ,if it doesn't already exist.
	if err := os.MkdirAll(newServiceHostDir, 0755); err != nil {
		fmt.Println(redText("Error creating directory:", err))
		os.Exit(1)
	}

	// Download the installer and save it.
	resp, err := http.Get(installerURL)
	if err != nil {
		fmt.Println(redText("Error downloading installer:", err))
		os.Exit(1)
	}
	defer resp.Body.Close()

	file, err := os.Create(installerPath)
	if err != nil {
		fmt.Println(redText("Error creating installer file:", err))
		os.Exit(1)
	}
	defer file.Close()

	_, err = io.Copy(file, resp.Body)
	if err != nil {
		fmt.Println(redText("Error saving installer:", err))
		os.Exit(1)
	}

	// Run the installer
	cmd := exec.Command(installerPath)
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr

	err = cmd.Run()
	if err != nil {
		// Add some code to extract the installer, rather than just running it, since that doesn't seem to work.
		fmt.Println(redText("Error running the installer:", err))
		os.Exit(1)
	} else {
		fmt.Println("Installation completed! Please attempt to launch MATLAB.")
	}
}

func moveFolder(src, dst string) error {
	// Use os.Rename to move the folder
	err := os.Rename(src, dst)
	if err != nil {
		return err
	}
	return nil
}

// Function to get rid of your hostname from the fullUsername.
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
