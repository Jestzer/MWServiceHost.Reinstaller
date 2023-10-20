package main

import (
	"bufio"
	"fmt"
	"os"
	"os/user"
	"strings"

	"github.com/shirou/gopsutil/process"
)

func main() {

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
					fmt.Printf("%s (PID %d) has been terminated.\n", name, p.Pid)
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
		fmt.Printf("%s' is not running. Skipping termination. \n", processName)
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

	dirs, err := os.ReadDir(searchDirectory)
	if err != nil {
		fmt.Printf("Error reading the directory: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("Folders starting with 'v':")
	for _, dir := range dirs {
		if dir.IsDir() && strings.HasPrefix(dir.Name(), "v") {
			fmt.Println(dir.Name())
		}
	}
}

// Get rid of your computer's name. We don't care about it.
func extractUsername(fullUsername string) string {
	parts := strings.SplitN(fullUsername, "\\", 2)
	if len(parts) == 2 {
		return parts[1]
	}
	return fullUsername
}
