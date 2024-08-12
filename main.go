package main

import (
	"archive/zip"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"runtime"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

type MyMainWindow struct {
	*walk.MainWindow
	filePathEdit   *walk.LineEdit
	outputPathEdit *walk.LineEdit
	progressBar    *walk.ProgressBar
	statusLabel    *walk.Label
	conversionInfo *walk.Label
}

const pandocVersion = "3.1.9"

func main() {
	ensurePandocInstalled()
	mw := &MyMainWindow{}

	icon, _ := walk.NewIconFromFile("icon.ico")

	if err := (MainWindow{
		AssignTo:   &mw.MainWindow,
		Title:      "Document Converter",
		MinSize:    Size{Width: 600, Height: 400},
		Size:       Size{Width: 600, Height: 400},
		Layout:     VBox{},
		Icon:       icon,
		Background: SolidColorBrush{Color: walk.RGB(240, 240, 240)},
		Children: []Widget{
			Composite{
				Layout: HBox{MarginsZero: true},
				Children: []Widget{
					HSpacer{},
					ImageView{
						Image: "logo.png",
						Mode:  ImageViewModeZoom,
					},
					HSpacer{},
				},
			},
			Composite{
				Layout: VBox{},
				Children: []Widget{
					Label{
						Text:      "Welcome to Document Converter",
						Font:      Font{Family: "Segoe UI", PointSize: 18, Bold: true},
						TextColor: walk.RGB(0, 120, 215),
					},
					Label{
						Text: "Convert your documents with ease",
						Font: Font{Family: "Segoe UI", PointSize: 12},
					},
				},
			},
			Composite{
				Layout: Grid{Columns: 3},
				Children: []Widget{
					Label{Text: "Input File:"},
					LineEdit{AssignTo: &mw.filePathEdit},
					PushButton{
						Text:      "Choose",
						OnClicked: mw.selectFile,
						MinSize:   Size{Width: 100},
					},
					Label{Text: "Output Folder:"},
					LineEdit{AssignTo: &mw.outputPathEdit, Text: getDefaultOutputDir()},
					PushButton{
						Text:      "Choose",
						OnClicked: mw.selectOutputFolder,
						MinSize:   Size{Width: 100},
					},
				},
			},
			Label{
				AssignTo: &mw.conversionInfo,
				Text:     "Conversion: [None Selected]",
				Font:     Font{PointSize: 12, Bold: true},
			},
			PushButton{
				Text:       "Start Conversion",
				OnClicked:  mw.convertFile,
				MinSize:    Size{Width: 150, Height: 40},
				Font:       Font{PointSize: 14, Bold: true},
				Background: SolidColorBrush{Color: walk.RGB(0, 120, 215)},
			},
			ProgressBar{
				AssignTo: &mw.progressBar,
				MinValue: 0,
				MaxValue: 100,
			},
			Label{
				AssignTo: &mw.statusLabel,
				Text:     "Status: Ready",
				Font:     Font{PointSize: 12},
			},
		},
	}.Create()); err != nil {
		fmt.Println("Error:", err)
		os.Exit(1)
	}

	mw.Run()
}

func getDefaultOutputDir() string {
	homeDir, _ := os.UserHomeDir()
	return filepath.Join(homeDir, "Downloads")
}

func (mw *MyMainWindow) selectFile() {
	dlg := new(walk.FileDialog)
	dlg.Title = "Choose file to convert"
	dlg.Filter = "Supported Files (*.md;*.docx;*.xlsx)|*.md;*.docx;*.xlsx"

	if ok, _ := dlg.ShowOpen(mw); !ok {
		return
	}
	mw.filePathEdit.SetText(dlg.FilePath)
	mw.updateConversionInfo()
}

func (mw *MyMainWindow) selectOutputFolder() {
	dlg := new(walk.FileDialog)
	dlg.Title = "Choose output folder"

	if ok, _ := dlg.ShowBrowseFolder(mw); !ok {
		return
	}
	mw.outputPathEdit.SetText(dlg.FilePath)
}

func (mw *MyMainWindow) updateConversionInfo() {
	filePath := mw.filePathEdit.Text()
	ext := strings.ToLower(filepath.Ext(filePath))
	var conversionText string

	switch ext {
	case ".md":
		conversionText = "Markdown to Word"
	case ".docx":
		conversionText = "Word to Markdown"
	case ".xlsx":
		conversionText = "Excel to Markdown"
	default:
		conversionText = "[None Selected]"
	}

	mw.conversionInfo.SetText(fmt.Sprintf("Conversion: %s", conversionText))
}

func (mw *MyMainWindow) convertFile() {
	inputPath := mw.filePathEdit.Text()
	outputDir := mw.outputPathEdit.Text()

	if inputPath == "" || outputDir == "" {
		walk.MsgBox(mw, "Error", "Please select input file and output folder", walk.MsgBoxIconError)
		return
	}

	go func() {
		mw.Synchronize(func() {
			mw.progressBar.SetValue(50)
			mw.statusLabel.SetText("Status: Converting...")
		})

		outputPath := filepath.Join(outputDir, changeExtension(filepath.Base(inputPath)))
		err := convertWithLocalPandoc(inputPath, outputPath)

		mw.Synchronize(func() {
			if err != nil {
				errorMsg := fmt.Sprintf("Error converting file: %v", err)
				if strings.Contains(err.Error(), "pandoc not found") {
					errorMsg = "Pandoc not found. Please ensure the 'pandoc' folder is in the same directory as the application."
				}
				walk.MsgBox(mw, "Conversion Error", errorMsg, walk.MsgBoxIconError)
				mw.statusLabel.SetText("Status: Conversion Failed")
			} else {
				mw.statusLabel.SetText("Status: Conversion Complete")
			}
			mw.progressBar.SetValue(100)
		})
	}()
}

func convertWithLocalPandoc(inputPath, outputPath string) error {
	pandocPath, err := findPandoc()
	if err != nil {
		return fmt.Errorf("pandoc not found: %v", err)
	}

	cmd := exec.Command(pandocPath, inputPath, "-o", outputPath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("pandoc execution failed: %v\nOutput: %s", err, string(output))
	}
	return nil
}

func ensurePandocInstalled() {
	pandocPath, err := findPandoc()
	if err == nil {
		fmt.Println("Pandoc found at:", pandocPath)
		return
	}

	fmt.Println("Pandoc not found. Downloading and installing...")

	// Определяем URL для скачивания в зависимости от операционной системы
	var url string
	switch runtime.GOOS {
	case "windows":
		url = fmt.Sprintf("https://github.com/jgm/pandoc/releases/download/%s/pandoc-%s-windows-x86_64.zip", pandocVersion, pandocVersion)
	case "darwin":
		url = fmt.Sprintf("https://github.com/jgm/pandoc/releases/download/%s/pandoc-%s-macOS.zip", pandocVersion, pandocVersion)
	case "linux":
		url = fmt.Sprintf("https://github.com/jgm/pandoc/releases/download/%s/pandoc-%s-linux-amd64.tar.gz", pandocVersion, pandocVersion)
	default:
		fmt.Println("Unsupported operating system")
		os.Exit(1)
	}

	// Скачиваем архив
	resp, err := http.Get(url)
	if err != nil {
		fmt.Println("Error downloading Pandoc:", err)
		os.Exit(1)
	}
	defer resp.Body.Close()

	// Создаем временный файл для архива
	tmpFile, err := os.CreateTemp("", "pandoc-*.zip")
	if err != nil {
		fmt.Println("Error creating temp file:", err)
		os.Exit(1)
	}
	defer os.Remove(tmpFile.Name())

	// Сохраняем архив
	_, err = io.Copy(tmpFile, resp.Body)
	if err != nil {
		fmt.Println("Error saving Pandoc archive:", err)
		os.Exit(1)
	}

	// Распаковываем архив
	err = unzip(tmpFile.Name(), "pandoc")
	if err != nil {
		fmt.Println("Error extracting Pandoc:", err)
		os.Exit(1)
	}

	fmt.Println("Pandoc installed successfully")
}

func unzip(src, dest string) error {
	r, err := zip.OpenReader(src)
	if err != nil {
		return err
	}
	defer r.Close()

	os.MkdirAll(dest, 0755)

	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return err
		}
		defer rc.Close()

		path := filepath.Join(dest, f.Name)
		if f.FileInfo().IsDir() {
			os.MkdirAll(path, f.Mode())
		} else {
			os.MkdirAll(filepath.Dir(path), f.Mode())
			f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
			if err != nil {
				return err
			}
			defer f.Close()

			_, err = io.Copy(f, rc)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func findPandoc() (string, error) {
	possiblePaths := []string{
		"pandoc\\pandoc.exe",
		"pandoc.exe",
		filepath.Join("..", "pandoc", "pandoc.exe"),
		filepath.Join("pandoc", fmt.Sprintf("pandoc-%s", pandocVersion), "pandoc.exe"), // Добавляем новый путь
	}

	execPath, err := os.Executable()
	if err == nil {
		execDir := filepath.Dir(execPath)
		possiblePaths = append(possiblePaths,
			filepath.Join(execDir, "pandoc", "pandoc.exe"),
			filepath.Join(execDir, "pandoc.exe"),
			filepath.Join(execDir, "pandoc", fmt.Sprintf("pandoc-%s", pandocVersion), "pandoc.exe"), // Добавляем новый путь
		)
	}

	for _, path := range possiblePaths {
		if _, err := os.Stat(path); err == nil {
			return path, nil
		}
	}

	return "", fmt.Errorf("pandoc not found in any of the expected locations")
}

func changeExtension(filename string) string {
	ext := filepath.Ext(filename)
	basename := filename[:len(filename)-len(ext)]

	switch ext {
	case ".md":
		return basename + ".docx"
	case ".docx", ".xlsx":
		return basename + ".md"
	default:
		return filename
	}
}