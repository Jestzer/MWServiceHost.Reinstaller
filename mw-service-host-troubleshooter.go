package main

import (
	"archive/zip"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"os/signal"
	"os/user"
	"path/filepath"
	"runtime"
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
	var serviceHostFolderExists bool
	var matlabConnectorFolderExists bool

	// Don't hate me.
	switch userOS := runtime.GOOS; userOS {
	case "darwin":
		fmt.Println(redText("Sorry, macOS is currently unsupported. :("))
		os.Exit(1)
	case "windows":
		// scrub
	case "linux":
		fmt.Println(redText("Sorry, Linux is currently unsupported. :("))
		os.Exit(1)
	default:
		fmt.Print(redText("\nYour operating system is unrecognized. Exiting."))
		os.Exit(1)
	}

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
		errorMsg := redText(fmt.Sprintf("Error: %v\n", err))
		fmt.Print(errorMsg)
		os.Exit(1)
	}

	// Loop through to look the list of processes and check if MWSH is running.
	for _, p := range processes {
		name, err := p.Name()
		if err == nil && strings.EqualFold(name, processName) {
			processDetected = true
			fmt.Printf("Process '%s' is running with PID %d\n", name, p.Pid)

			// Prompt the user to end the process.
			fmt.Print("Do you want to end this process? (y/n)\n")

			for {
				input, err := rl.Readline()
				if err != nil {
					if err.Error() == "Interrupt" {
						fmt.Println(redText("Exiting from user input."))
					} else {
						fmt.Print(redText("Error reading line: ", err))
						continue
					}
					return
				}
				input = strings.TrimSpace(input)
				input = strings.ToLower(input)

				if input == "y" || input == "yes" {
					err := p.Terminate()
					if err == nil {
						fmt.Printf("%s (PID %d) has been terminated.\nPlease wait for the next step.", name, p.Pid)
					} else {
						errorMsg := redText(fmt.Sprintf("Error terminating the process: %v\n", err))
						fmt.Print(errorMsg)
						os.Exit(1)
					}
					break
				} else if input == "n" || input == "no" {
					fmt.Println(redText("Exiting since the user has declined moving necessary folders."))
					os.Exit(0)
				} else {
					fmt.Println(redText("Invalid input. Type in \"y\" or \"n\""))
					continue
				}
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
		errorMsg := redText(fmt.Sprintf("Error detecting the Windows username: %v\n", err))
		fmt.Print(errorMsg)
		os.Exit(1)
	}
	username := extractUsername(currentUser.Username)

	// Look for any installed ServiceHosts by searching for "v" folders.
	serviceHostFolder := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks\\ServiceHost", username)

	// Define the Local MathWorks folder/directory.
	matlabConnectorFolder := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\MathWorks\\MATLABConnector", username)

	// Define the new Service Host download directory.
	newServiceHostDownloadFolder := fmt.Sprintf("C:\\Users\\%s\\AppData\\Local\\Temp", username)

	// Define where the old, moved Service Host folders will go.
	movedServiceHostFilesFolder := filepath.Join(thisProgramDir, "Old_service_host_files")

	// Define where the Service Host extracted installation files will go.
	extractedServiceHostInstaller := filepath.Join(newServiceHostDownloadFolder, "extractedMWSInstaller")

	// Check if the existing Service Host folder exists.
	if _, err := os.Stat(serviceHostFolder); err == nil {
		serviceHostFolderExists = true
	} else {
		serviceHostFolderExists = false
	}

	// Check if the existing MATLAB Connector folder exists.
	if _, err := os.Stat(matlabConnectorFolder); err == nil {
		matlabConnectorFolderExists = true
	} else {
		matlabConnectorFolderExists = false
	}

	if serviceHostFolderExists || matlabConnectorFolderExists {
		// Ask the user if they are ready to move the old ServiceHost folders.
		fmt.Println("\nAre you ready to move the old ServiceHost folder, MATLAB Connector folder, their contents, and all other related files?")
		fmt.Print(cyanText("Any related content and folders marked for being moved above will be moved into a folder called \"Old_service_host_files\"\n"))
		fmt.Print(cyanText(". You are responsible for deleting this folder (if you wish) after the old Service Host files are moved into this folder."))
		fmt.Print(cyanText("\nType \"move\" to confirm. This is case-sensitive. Type any anything else to cancel.\n"))

		for {
			input, err := rl.Readline()
			if err != nil {
				if err.Error() == "Interrupt" {
					fmt.Println(redText("Exiting from user input."))
				} else {
					fmt.Print(redText("Error reading line: ", err))
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

		// Create the junk folder, if it doesn't already exist. Add a number to the end of the folder if you've already tried this.
		// However, if it's failing after 25 times, just assume something is wrong.
		i := 1
		for {
			movedServiceHostFilesFolder = filepath.Join(thisProgramDir, "Old_service_host_files")
			if i > 1 {
				movedServiceHostFilesFolder = fmt.Sprintf("%s%d", movedServiceHostFilesFolder, i)
			}

			if _, err := os.Stat(movedServiceHostFilesFolder); os.IsNotExist(err) {
				// If the folder does not exist, try to create it.
				if err := os.MkdirAll(movedServiceHostFilesFolder, 0755); err != nil {
					fmt.Println(redText("Error creating directory: ", err))
					os.Exit(1)
				} else {
					fmt.Println("Directory created successfully:", movedServiceHostFilesFolder)
					break
				}
			} else if i >= 26 {
				// If the folder exists and we've reached the maximum attempt count, exit with an error.
				fmt.Println(redText("Error: Reached maximum directory creation attempts."))
				os.Exit(1)
			}

			i++
		}

		if serviceHostFolderExists {
			// Move the ServiceHost folder.
			err := moveFolder(serviceHostFolder, filepath.Join(movedServiceHostFilesFolder, "ServiceHost"))
			if err != nil {
				errorMsg := redText(fmt.Sprintf("Error moving ServiceHost folder: %s\n", err))
				fmt.Print(errorMsg)
				os.Exit(1)
			} else {
				fmt.Println("ServiceHost folder moved successfully.")
			}
		}
	}

	if matlabConnectorFolderExists {
		// Move the MATLABConnector folder.
		err = moveFolder(matlabConnectorFolder, filepath.Join(movedServiceHostFilesFolder, "MATLABConnector"))
		if err != nil {
			errorMsg := redText(fmt.Sprintf("Error moving MATLABConnector folder: %s\n", err))
			fmt.Print(errorMsg)
			os.Exit(1)
		} else {
			fmt.Println("MATLABConnector folder moved successfully.")
		}
	}

	fmt.Println("Downloading latest Service Host installer. Please wait.")

	installerURL := "https://ssd.mathworks.com/supportfiles/downloads/MathWorksServiceHost/v2024.2.0.3/installers/mathworksservicehost_2024.2.0.3_win64_installer.exe"
	installerPath := filepath.Join(newServiceHostDownloadFolder, "mathworksservicehost_installer.exe")

	// Check if the downloaeded Service Host WinZip executable already exists. Delete it if it is.
	if _, err := os.Stat(installerPath); err == nil {
		err := os.Remove(installerPath)
		if err != nil {
			fmt.Print(redText("\nFailed to delete the existing Service Host installer directory: ", err))
			os.Exit(1)
		}
	}

	// Download the installer and save it.
	resp, err := http.Get(installerURL)
	if err != nil {
		fmt.Println(redText("Error downloading installer: ", err))
		os.Exit(1)
	}
	defer resp.Body.Close()

	file, err := os.Create(installerPath)
	if err != nil {
		fmt.Println(redText("Error creating installer file: ", err))
		os.Exit(1)
	}
	defer file.Close()

	_, err = io.Copy(file, resp.Body)
	if err != nil {
		fmt.Println(redText("Error saving installer: ", err))
		os.Exit(1)
	}

	// Check if the extracted Service Host installer directory already exists. Delete it if it is.
	if _, err := os.Stat(extractedServiceHostInstaller); err == nil {
		err := os.RemoveAll(extractedServiceHostInstaller)
		if err != nil {
			fmt.Print(redText("\nFailed to delete the existing Service Host installer directory: ", err))
			os.Exit(1)
		}
	}

	// Extract the WinZIP archive.
	err = unzipFile(installerPath, extractedServiceHostInstaller)
	if err != nil {
		fmt.Print(redText("\nFailed to extract Service Host installer: ", err))
		os.Exit(1)
	}

	extractedInstallerPath := filepath.Join(extractedServiceHostInstaller, "bin\\win64\\InstallMathWorksServiceHost.exe")

	// Run the installer
	cmd := exec.Command(extractedInstallerPath)
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr

	err = cmd.Run()
	if err != nil {
		if strings.Contains(err.Error(), "exit status 66") {
			fmt.Println(cyanText("Installation started! Please wait.\n"))
			fmt.Print(cyanText("When you are prompted to \"Setup MATLAB Drive Connector\", you may then select\"Exit\" and then \"OK\""))
		} else {
			fmt.Println(redText("Error running the installer: ", err))
			os.Exit(1)
		}
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

// Function to unzip integration scripts.
func unzipFile(src, dest string) error {
	reader, err := zip.OpenReader(src)
	if err != nil {
		return err
	}
	defer reader.Close()

	for _, file := range reader.File {
		path := filepath.Join(dest, file.Name)

		// Reconstruct the file path on Windows to ensure proper subdirectories are created. Don't know why other OSes don't need this.
		if runtime.GOOS == "windows" {
			path = filepath.Join(dest, file.Name)
			path = filepath.FromSlash(path)
		}

		if file.FileInfo().IsDir() {
			os.MkdirAll(path, file.Mode())
			continue
		}

		err := os.MkdirAll(filepath.Dir(path), 0755)
		if err != nil {
			return err
		}

		fileReader, err := file.Open()
		if err != nil {
			return err
		}
		defer fileReader.Close()

		targetFile, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, file.Mode())
		if err != nil {
			return err
		}
		defer targetFile.Close()

		_, err = io.Copy(targetFile, fileReader)
		if err != nil {
			return err
		}
	}
	return nil
}
