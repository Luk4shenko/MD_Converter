package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/spreadsheet"
)

type MyMainWindow struct {
	*walk.MainWindow
	filePathEdit   *walk.LineEdit
	outputPathEdit *walk.LineEdit
	progressBar    *walk.ProgressBar
	statusLabel    *walk.Label
	conversionInfo *walk.Label
}

func main() {

	mw := &MyMainWindow{}

	icon, _ := walk.NewIconFromFile("icon.ico")

	if err := (MainWindow{
		AssignTo:  &mw.MainWindow,
		Title:     "Document Converter",
		MinSize:   Size{Width: 600, Height: 400},
		Size:      Size{Width: 600, Height: 400},
		Layout:    VBox{},
		Icon:      icon,
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
					Label{Text: "Input Files:"},
					LineEdit{AssignTo: &mw.filePathEdit},
					PushButton{
						Text:      "Choose",
						OnClicked: mw.selectFiles,
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
				Text:      "Start Conversion",
				OnClicked: mw.convertFiles,
				MinSize:   Size{Width: 150, Height: 40},
				Font:      Font{PointSize: 14, Bold: true},
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

func (mw *MyMainWindow) selectFiles() {
	dlg := new(walk.FileDialog)
	dlg.Title = "Choose files to convert"
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
	mw.updateConversionInfo()
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

func (mw *MyMainWindow) convertFiles() {
	filePath := mw.filePathEdit.Text()
	outputDir := mw.outputPathEdit.Text()

	if filePath == "" || outputDir == "" {
		walk.MsgBox(mw, "Error", "Please select input file and output folder", walk.MsgBoxIconError)
		return
	}

	mw.progressBar.SetRange(0, 100)
	mw.progressBar.SetValue(0)
	mw.statusLabel.SetText("Status: Converting...")

	go func() {
		err := convertFile(filePath, outputDir, func(progress int) {
			mw.Synchronize(func() {
				mw.progressBar.SetValue(progress)
			})
		})

		mw.Synchronize(func() {
			if err != nil {
				walk.MsgBox(mw, "Conversion Error", fmt.Sprintf("Error converting file: %v", err), walk.MsgBoxIconError)
				mw.statusLabel.SetText("Status: Conversion Failed")
			} else {
				mw.statusLabel.SetText("Status: Conversion Complete")
			}
			mw.progressBar.SetValue(100)
		})
	}()
}

func convertFile(filePath, outputDir string, progressCallback func(int)) error {
	ext := strings.ToLower(filepath.Ext(filePath))
	switch ext {
	case ".md":
		return convertMarkdownToWord(filePath, outputDir, progressCallback)
	case ".docx":
		return convertWordToMarkdown(filePath, outputDir, progressCallback)
	case ".xlsx":
		return convertExcelToMarkdown(filePath, outputDir, progressCallback)
	}
	return fmt.Errorf("unsupported file type: %s", ext)
}

func convertMarkdownToWord(filePath, outputDir string, progressCallback func(int)) error {
	content, err := ioutil.ReadFile(filePath)
	if err != nil {
		return err
	}

	doc := document.New()
	lines := strings.Split(string(content), "\n")

	listLevel := 0
	inList := false

	for i, line := range lines {
		progressCallback(int(float64(i) / float64(len(lines)) * 100))

		line = strings.TrimSpace(line)

		if line == "" {
			if !inList {
				doc.AddParagraph()
			}
			continue
		}

		if strings.HasPrefix(line, "#") {
			level := strings.Count(strings.Split(line, " ")[0], "#")
			text := strings.TrimSpace(line[level:])
			para := doc.AddParagraph()
			para.Properties().SetStyle(fmt.Sprintf("Heading%d", level))
			run := para.AddRun()
			run.AddText(text)
			inList = false
		} else if strings.HasPrefix(line, "- ") || strings.HasPrefix(line, "* ") || strings.HasPrefix(line, "+ ") {
			para := doc.AddParagraph()
			para.SetStyle("ListParagraph")
			para.SetNumberingLevel(listLevel)
			run := para.AddRun()
			run.AddText(strings.TrimPrefix(strings.TrimPrefix(strings.TrimPrefix(line, "- "), "*"), "+"))
			inList = true
		} else if len(line) > 0 && line[0] >= '0' && line[0] <= '9' && strings.Contains(line, ".") {
			para := doc.AddParagraph()
			para.SetStyle("ListParagraph")
			para.SetNumberingLevel(listLevel)
			run := para.AddRun()
			run.AddText(strings.TrimSpace(strings.SplitN(line, ".", 2)[1]))
			inList = true
		} else {
			para := doc.AddParagraph()
			run := para.AddRun()
			run.AddText(line)
			inList = false
		}
	}

	outputPath := filepath.Join(outputDir, strings.TrimSuffix(filepath.Base(filePath), ".md")+".docx")
	err = doc.SaveToFile(outputPath)
	progressCallback(100)
	return err
}

func convertWordToMarkdown(filePath, outputDir string, progressCallback func(int)) error {
    doc, err := document.Open(filePath)
    if err != nil {
        return err
    }

    var markdown strings.Builder
    totalParagraphs := len(doc.Paragraphs())

    for i, para := range doc.Paragraphs() {
        progressCallback(int(float64(i) / float64(totalParagraphs) * 100))

        text := ""
        for _, run := range para.Runs() {
            if run.Properties().IsBold() {
                text += "**" + run.Text() + "**"
            } else if run.Properties().IsItalic() {
                text += "_" + run.Text() + "_"
            } else {
                text += run.Text()
            }
        }
        
        style := para.Style()
        if style != "" {
            switch style {
            case "Heading1":
                markdown.WriteString("# " + text + "\n\n")
            case "Heading2":
                markdown.WriteString("## " + text + "\n\n")
            case "Heading3":
                markdown.WriteString("### " + text + "\n\n")
            case "ListParagraph":
                // Check if the paragraph has numbering
                if isNumberedParagraph(para) {
                    level := getNumberingLevel(para)
                    indent := strings.Repeat("  ", level)
                    markdown.WriteString(indent + "1. " + text + "\n")
                } else {
                    markdown.WriteString("- " + text + "\n")
                }
            default:
                markdown.WriteString(text + "\n\n")
            }
        } else {
            markdown.WriteString(text + "\n\n")
        }
    }

    outputPath := filepath.Join(outputDir, strings.TrimSuffix(filepath.Base(filePath), ".docx")+".md")
    err = ioutil.WriteFile(outputPath, []byte(markdown.String()), 0644)
    progressCallback(100)
    return err
}

func isNumberedParagraph(para document.Paragraph) bool {
    props := para.Properties()
    // Check if the paragraph has any numbering properties
    return props.X().NumPr != nil
}

func getNumberingLevel(para document.Paragraph) int {
    props := para.Properties()
    if props.X().NumPr == nil {
        return 0
    }
    // Try to get the numbering level
    if props.X().NumPr.Ilvl != nil {
        return int(props.X().NumPr.Ilvl.ValAttr)
    }
    return 0
}

func convertExcelToMarkdown(filePath, outputDir string, progressCallback func(int)) error {
	wb, err := spreadsheet.Open(filePath)
	if err != nil {
		return err
	}

	var markdown strings.Builder
	totalSheets := len(wb.Sheets())

	for sheetIndex, sheet := range wb.Sheets() {
		markdown.WriteString("# " + sheet.Name() + "\n\n")

		rows := sheet.Rows()
		for rowIndex, row := range rows {
			progressCallback(int(float64(sheetIndex*len(rows)+rowIndex) / float64(totalSheets*len(rows)) * 100))

			if len(row.Cells()) == 0 {
				continue
			}

			cells := row.Cells()
			if len(cells) == 0 {
				continue
			}

			if rowIndex == 0 {
				markdown.WriteString("|")
				for _, cell := range cells {
					markdown.WriteString(" " + cell.GetString() + " |")
				}
				markdown.WriteString("\n|")
				for range cells {
					markdown.WriteString(" --- |")
				}
				markdown.WriteString("\n")
			} else {
				markdown.WriteString("|")
				for _, cell := range cells {
					markdown.WriteString(" " + cell.GetString() + " |")
				}
				markdown.WriteString("\n")
			}
		}

		markdown.WriteString("\n")
	}

	outputPath := filepath.Join(outputDir, strings.TrimSuffix(filepath.Base(filePath), ".xlsx")+".md")
	err = ioutil.WriteFile(outputPath, []byte(markdown.String()), 0644)
	progressCallback(100)
	return err
}