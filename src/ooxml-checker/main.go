package main

import (
	"encoding/base64"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strings"

	flags "github.com/jessevdk/go-flags"
	color "github.com/labstack/gommon/color"
	"github.com/richardlehane/mscfb"
	util "github.com/woanware/goutil"
)

// ##### Structs ##############################################################

type Options struct {
	Input   string `short:"i" long:"input" description:"Input file" required:"true"`
	Output  string `short:"o" long:"output" description:"Output directory" required:"true"`
	NoUnzip bool   `short:"n" long:"nonzip" description:"Don't unzip file, just process" required:"false"`
}

var (
	options Options

	reImage      = regexp.MustCompile(`\sType="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image"\s*?Target="(.*?)"/>`)
	reAudio      = regexp.MustCompile(`\sType="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/audio"\s*?Target="(.*?)"/>`)
	reHyperLink  = regexp.MustCompile(`\sType="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"\s*?Target="(.*?)"/>`)
	reVba        = regexp.MustCompile(`(?m)\sType="http://schemas\.microsoft\.com/office/2006/relationships/vbaProject"\s*?Target="(.*?)"/>`)
	reBinaryData = regexp.MustCompile(`<pkg:part\spkg:name="([\/a-z\.A-Z0-9]*)"\spkg:contentType="application/vnd.ms-office.vbaProject"><pkg:binaryData>([a-zA-Z0-9/\n\r+=]*?)</pkg:binaryData>`)
	reMacro      = regexp.MustCompile(`<wne:mcd wne:macroName="(.*?)" wne:name="(.*?)"\swne:bEncrypt="(.*?)"`)
	reDde1       = regexp.MustCompile(`ddeService="(.*?)"\sddeTopic="(.*?)">`)
	reDde2       = regexp.MustCompile(`<w:instrText>(.*?DDEAUTO\s.*?)/w:instrText>`)
	reDde3       = regexp.MustCompile(`<w:instrText>(.*?DDE\s.*?)/w:instrText>`)
)

// ##### Constants #####################################################################################################

const APP_NAME string = "ooxml-checker"
const APP_VERSION string = "0.0.1"

const (
	OOXML_REGEX_IMAGE       int = 0
	OOXML_REGEX_AUDIO       int = 1
	OOXML_REGEX_HYPERLINK   int = 2
	OOXML_REGEX_VBA         int = 3
	OOXML_REGEX_BINARY_DATA int = 4
	OOXML_REGEX_MACRO       int = 5
	OOXML_REGEX_DDE1        int = 6
	OOXML_REGEX_DDE2        int = 7
	OOXML_REGEX_DDE3        int = 8
)

// ##### Methods ##############################################################

// main is the application entry point!
func main() {

	fmt.Println(fmt.Sprintf("\n%s v%s - woanware\n", APP_NAME, APP_VERSION))

	parseCommandLine()

	if options.Output != "." {
		if util.DoesDirExist(options.Output) == false {
			os.MkdirAll(options.Output, 0700)
		}
	}

	regexesString := make([]*regexp.Regexp, 0)
	regexesString = append(regexesString, reImage)
	regexesString = append(regexesString, reAudio)
	regexesString = append(regexesString, reHyperLink)
	regexesString = append(regexesString, reVba)
	regexesString = append(regexesString, reBinaryData)
	regexesString = append(regexesString, reMacro)
	regexesString = append(regexesString, reDde1)
	regexesString = append(regexesString, reDde2)
	regexesString = append(regexesString, reDde3)

	var files []string
	var err error
	if options.NoUnzip == true {
		files = append(files, options.Input)
	} else {
		files, err = util.Unzip(options.Input, strings.TrimSuffix(options.Input, path.Ext(options.Input)))
		if err != nil {
			fmt.Printf("Checking unzipping file: %v\n", err)
			return
		}
	}

	for _, file := range files {
		checkFile(files, options.Output, file, regexesString)
	}
}

// parseCommandLine parses the command line options
func parseCommandLine() {

	var parser = flags.NewParser(&options, flags.Default)
	if _, err := parser.Parse(); err != nil {
		if flagsErr, ok := err.(*flags.Error); ok && flagsErr.Type == flags.ErrHelp {
			os.Exit(0)
		} else {
			os.Exit(1)
		}
	}
}

// checkFile checks each file for the data relationships by calling the various regexes
func checkFile(files []string, outputPath string, filePath string, regexesString []*regexp.Regexp) {

	fmt.Printf("Checking file %s\n", filePath)

	data, _ := ioutil.ReadFile(filePath)

	for i, regex := range regexesString {
		for _, match := range regex.FindAllStringSubmatch(string(data), -1) {

			switch i {
			case OOXML_REGEX_IMAGE:
				color.Println(color.Red(fmt.Sprintf("Found image data: %s", match[1])))

			case OOXML_REGEX_AUDIO:
				color.Println(color.Red(fmt.Sprintf("Found audio data: %s", match[1])))

			case OOXML_REGEX_HYPERLINK:
				color.Println(color.Red(fmt.Sprintf("Found hyperlink data: %s", match[1])))

			case OOXML_REGEX_VBA:
				color.Println(color.Red(fmt.Sprintf("Found vbaProject data: %s", match[1])))
				processOle(files, match[1])

			case OOXML_REGEX_BINARY_DATA:
				color.Println(color.Red(fmt.Sprintf("Found vbaProject binary data: %s", match[1])))
				processBinaryData(match[1], match[2])

			case OOXML_REGEX_MACRO:
				color.Println(color.Red(fmt.Sprintf("Found macro data: Name: %s # MacroName: %s # bEncrypt: %v", match[1], match[2], match[3])))

			case OOXML_REGEX_DDE1:
				color.Println(color.Red(fmt.Sprintf("Found DDE data: Service: %s # Topic: %s", match[1], match[2])))

			case OOXML_REGEX_DDE2, OOXML_REGEX_DDE3:
				color.Println(color.Red(fmt.Sprintf("Found DDE data: %s", match[1])))
			}
		}
	}
}

// processOle finds the OLE file needed for processing, then
// calls the function to parse and extract the OLE data
func processOle(files []string, fileName string) {

	for _, file := range files {
		// Lets try and find a file that matches the part previously identified
		if filepath.Base(file) != fileName {
			continue
		}

		extractOle(file)
	}
}

// extractOle parses the OLE data and extracts each element to a separate file
func extractOle(filePath string) {

	file, err := os.Open(filePath)
	if err != nil {
		fmt.Printf("Error opening OLE data: %v", err)
		return
	}
	defer file.Close()

	doc, err := mscfb.New(file)
	if err != nil {
		fmt.Printf("Error parsing OLE data: %v", err)
		return
	}

	fileName := ""
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		buf := make([]byte, 512)
		i, _ := doc.Read(buf)
		if i > 0 {
			fileName = filepath.Join(options.Output, "ole_"+util.RemoveIllegalPathCharacters(entry.Name))
			err = ioutil.WriteFile(fileName, buf[:i], 0700)
			if err != nil {
				fmt.Printf("Error writing OLE data: %v", err)
				continue
			}

			color.Println(color.Yellow(fmt.Sprintf("Wrote OLE file: %s", fileName)))
		}

		color.Println(color.Red(fmt.Sprintf("Found OLE entry: %s", entry.Name)))
	}
}

// processBinaryData base64 decodes the binary data, writes
//  it to a file, and calls the OLE parsing functionality
func processBinaryData(partName string, data string) {

	temp, err := base64.StdEncoding.DecodeString(data)
	if err != nil {
		fmt.Printf("Error base64 decoding binary data: %v", err)
		return
	}

	fileName := filepath.Join(options.Output, util.RemoveIllegalPathCharacters(partName))

	err = ioutil.WriteFile(fileName, temp, 0700)
	if err != nil {
		fmt.Printf("Error writing base64 decoded binary data: %v", err)
		return
	}

	color.Println(color.Yellow(fmt.Sprintf("Wrote Base64 decoded binary file: %s", fileName)))

	extractOle(fileName)
}
