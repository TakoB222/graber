package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	//"sync"

	excel "github.com/360EntSecGroup-Skylar/excelize"
	"github.com/PuerkitoBio/goquery"
)

var (
	fileDatesRe = regexp.MustCompile(`_(\d+)-(\d+)\.htm`)
)

func main() {
	var (
		dataPath  = "."
		excelFile = excel.NewFile()
	)

	if len(os.Args) > 1 {
		dataPath = os.Args[1]
	}

	count := 0
	dirNames := GrabDir(dataPath)

	for _, dirName := range dirNames {
		fileNames := GrabFiles(dirName)
		count += WriteReport(excelFile, filepath.Base(dirName), fileNames)
	}

	if count > 0 {
		err := excelFile.SaveAs(filepath.Join(dataPath, "all.xlsx"))
		if err != nil {
			log.Fatalln(err)
		}
	}
}

func getFileDates(fileName string) (from, till time.Time) {
	submatch := fileDatesRe.FindAllStringSubmatch(fileName, -1)
	if len(submatch) > 0 {
		if len(submatch[0]) == 3 {
			from, _ = time.Parse("20060102", submatch[0][1])
			till, _ = time.Parse("20060102", submatch[0][2])
		}
	}
	return
}

func GrabDir(dataPath string) (dirNames []string) {
	dataPath = filepath.Join(dataPath, "*")
	dirNames, _ = filepath.Glob(dataPath)
	tmp := dirNames[:0]
	for _, dirName := range dirNames {
		if stat, err := os.Stat(dirName); err == nil {
			if stat.IsDir() {
				tmp = append(tmp, dirName)
			}
		}
	}
	dirNames = tmp
	sort.Strings(dirNames)
	return
}

func GrabFiles(dataPath string) (fileNames []string) {
	dataPath = filepath.Join(dataPath, "*")
	fileNames, _ = filepath.Glob(dataPath)
	tmp := fileNames[:0]
	for _, fileName := range fileNames {
		switch {
		case strings.HasSuffix(fileName, ".html"), strings.HasSuffix(fileName, ".htm"):
			tmp = append(tmp, fileName)
		}
	}
	fileNames = tmp

	sort.Slice(fileNames, func(i, j int) bool {
		from1, _ := getFileDates(fileNames[i])
		from2, _ := getFileDates(fileNames[j])
		return from1.After(from2)
	})

	return
}

func WriteReport(excelFile *excel.File, sheetName string, fileNames []string) int {

	fmt.Println("WriteReport:", sheetName)

	sheet := excelFile.NewSheet(sheetName)
	excelFile.SetActiveSheet(sheet)

	excelFile.DeleteSheet("Sheet1")

	offset := 0

	for _, fileName := range fileNames {
		var rows = ParseReport(fileName)

		fmt.Println("", fileName)

		if len(rows) > 0 {
			var from, till = getFileDates(fileName)


			WriteSheet(excelFile, sheetName, offset, from, till, rows)
			offset++
		}
	}

	return offset
}

func WriteSheet(excelFile *excel.File, sheetName string, offset int, from, till time.Time, rows []ReportRow) {
	var coll = 'A' + 4*offset

	coords := func(x, y int) string {
		return fmt.Sprintf("%c%d", coll+x, y)
	}

	excelFile.SetCellValue(sheetName, coords(0, 1), from.Format("2006-01-02"))

	excelFile.SetCellValue(sheetName, coords(0, 2), "Проход")
	excelFile.SetCellValue(sheetName, coords(1, 2), "Прибыль")
	excelFile.SetCellValue(sheetName, coords(2, 2), "Всего сделок")
	excelFile.SetCellValue(sheetName, coords(3, 2), "Просадка $")

	for n, row := range rows {
		pass, _ := strconv.Atoi(row.Pass)
		profit, _ := strconv.ParseFloat(row.Profit, 64)
		drawDown, _ := strconv.ParseFloat(row.DrawDown, 64)
		totalTrades, _ := strconv.Atoi(row.TotalTrades)

		excelFile.SetCellValue(sheetName, coords(0, 3+n), pass)
		excelFile.SetCellValue(sheetName, coords(1, 3+n), profit)
		excelFile.SetCellValue(sheetName, coords(2, 3+n), totalTrades)
		excelFile.SetCellValue(sheetName, coords(3, 3+n), drawDown)
	}
}

type ReportRow struct {
	Pass        string
	Profit      string
	TotalTrades string
	DrawDown    string
}

func ParseReport(fileName string) (rows []ReportRow) {
	file, err := os.Open(fileName)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	doc, err := goquery.NewDocumentFromReader(file)
	if err != nil {
		log.Fatal(err)
	}

	doc.Find("table").Each(func(index int, tablehtml *goquery.Selection) {
		if index != 1 {
			return
		}

		tablehtml.Find("tr").Each(func(index int, rowhtml *goquery.Selection) {
			if index == 0 {
				return
			}

			var row ReportRow

			rowhtml.Find("td").Each(func(index int, tablecell *goquery.Selection) {
				text := strings.TrimSpace(tablecell.Text())

				switch index {
				case 0:
					row.Pass = text
				case 1:
					row.Profit = text
				case 2:
					row.TotalTrades = text
				case 5:
					row.DrawDown = text
				}
			})

			rows = append(rows, row)
		})
	})

	return
}
