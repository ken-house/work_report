package controller

import (
	"errors"
	"fmt"
	"github.com/shakinm/xlsReader/xls"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strconv"
	"sync"
	"tapd/work_report/xlsx/config"
	"time"
)

// XlsxController 操作xlsx文件对象
type XlsxController struct {
	dataController *DataController
}

// sheet名称
var globalSheetName string

var wg sync.WaitGroup
var lock sync.Mutex

// ParseXlsxOrXls 解析xlsx或xls文件
// @author ken
func (x *XlsxController) ParseXlsxOrXls(inputDirPath string) (dataArr []config.XlsxData, err error) {
	dataArr = make([]config.XlsxData, 0, 100)
	fileController := new(FileController)
	err = fileController.ScanDirPath(inputDirPath)
	if err != nil {
		fmt.Printf("ScanDirPath err:%v", err)
		return dataArr, err
	}
	if len(filePathArr) == 0 {
		return dataArr, errors.New("目录下没有文件")
	}

	// 使用goroutine并发处理，解析文件
	for _, filePath := range filePathArr {
		wg.Add(1)
		go func(f string) {
			var data config.XlsxData
			defer wg.Done()
			// 解析文件
			ext := filepath.Ext(f)
			if ext == ".xls" {
				data, err = x.ParseSignalXls(f)
			} else {
				data, err = x.ParseSignalXlsx(f)
			}
			if err != nil {
				fmt.Printf(f+",ParseXlsxOrXls err:%v\n", err)
				return
			}
			lock.Lock()
			dataArr = append(dataArr, data)
			lock.Unlock()
		}(filePath)
	}
	wg.Wait()
	if len(dataArr) == 0 {
		return dataArr, errors.New("文件没有数据")
	}
	return dataArr, nil
}

// ParseSignalXls 解析xls文件
// @author ken
func (x *XlsxController) ParseSignalXls(filePath string) (data config.XlsxData, err error) {
	wb, err := xls.OpenFile(filePath)
	if err != nil {
		return
	}

	rows := make([][]string, 0, 100)
	// 获取第一个sheet上所有单元格
	s, err := wb.GetSheet(0)
	if err != nil {
		return
	}

	for i := 0; i <= s.GetNumberRows(); i++ {
		cells, _ := s.GetRow(i)
		row := make([]string, 0, 30)
		for k := range cells.GetCols() {
			c, _ := cells.GetCol(k)
			row = append(row, c.GetString())
		}
		rows = append(rows, row)
	}
	return x.dataController.FormatDataFromXlsx(filePath, rows)
}

// Xls2Xlsx 将xls文件转为xlsx文件(项目中暂未使用)
// @author ken
func (x *XlsxController) Xls2Xlsx(filePath string) (savePath string, err error) {
	wb, err := xls.OpenFile(filePath)
	if err != nil {
		return
	}

	// 获取第一个sheet上所有单元格
	s, err := wb.GetSheet(0)
	if err != nil {
		return
	}

	savePath = filePath + "x"
	xlsxSheetName := s.GetName()
	xlsxFile, err := newOrOpenXlsxFile(savePath, xlsxSheetName)
	for i := 0; i <= s.GetNumberRows(); i++ {
		cells, _ := s.GetRow(i)
		row := make([]string, 0, 30)
		for k := range cells.GetCols() {
			c, _ := cells.GetCol(k)
			row = append(row, c.GetString())
		}
		err = xlsxFile.SetSheetRow(xlsxSheetName, fmt.Sprintf("A%d", i+1), &row)
		if err != nil {
			return
		}
	}

	err = xlsxFile.SaveAs(filePath + "x")
	if err != nil {
		return
	}

	// 删除原文件
	err = os.Remove(filePath)
	return
}

// ParseSignalXlsx 解析单个xlsx文件，读取文件内容
// @author ken
func (x *XlsxController) ParseSignalXlsx(filePath string) (data config.XlsxData, err error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
			return
		}
	}()

	// 获取第一个sheet上所有单元格
	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		fmt.Println(err)
		return
	}

	return x.dataController.FormatDataFromXlsx(filePath, rows)
}

// 检查文件是否存在，如果存在，则打开文件新增一个sheet；如果不存在，则创建一个工作簿
func newOrOpenXlsxFile(savePath string, sheetName string) (f *excelize.File, err error) {
	if existFile(savePath) {
		f, err = excelize.OpenFile(savePath)
		if err != nil {
			panic(err)
		}
		// 删除sheetName
		f.DeleteSheet(sheetName)
	} else {
		// 创建目录
		err = os.MkdirAll(filepath.Dir(savePath), 0755)
		if err != nil {
			panic(err)
		}
		f = excelize.NewFile()
	}
	defer func() {
		if err := f.Close(); err != nil {
			panic(err)
		}
	}()
	// 增加一个sheet
	index := f.NewSheet(sheetName)
	// 设置为当前sheet
	f.SetActiveSheet(index)
	// 删除默认的sheet1
	f.DeleteSheet("sheet1")
	return f, nil
}

// ExportXlsx 生成xlsx文件
func (x *XlsxController) ExportXlsx(savePath string, projectList []config.ProjectList) error {
	globalSheetName = time.Now().Format("1月2日")
	f, err := newOrOpenXlsxFile(savePath, globalSheetName)
	if err != nil {
		fmt.Printf("newOrOpenXlsxFile err:%v\n", err)
		return err
	}
	
	// 设置行宽
	f.SetColWidth(globalSheetName, "A", "A", 32.00)
	f.SetColWidth(globalSheetName, "B", "B", 10.00)
	f.SetColWidth(globalSheetName, "C", "C", 32.00)
	f.SetColWidth(globalSheetName, "D", "D", 15.00)
	f.SetColWidth(globalSheetName, "E", "E", 15.00)
	// 设置行高
	f.SetRowHeight(globalSheetName, 1, 26.00)
	f.SetRowHeight(globalSheetName, 2, 24.00)

	// 设置第一行
	err = setExcelFirstRow(f)
	if err != nil {
		fmt.Printf("setExcelFirstRow err:%v\n", err)
		return err
	}

	// 设置header行
	err = setExcelHeader(f)
	if err != nil {
		fmt.Printf("setExcelHeader err:%v\n", err)
		return err
	}

	// 设置项目成员工作内容
	lineNum, err := setTableData(f, projectList)
	if err != nil {
		fmt.Printf("setTableData err:%v\n", err)
		return err
	}

	// 设置末尾行
	err = setExcelEndRow(f, lineNum)
	if err != nil {
		fmt.Printf("setExcelEndRow err:%v\n", err)
		return err
	}

	// 保存文件
	err = f.SaveAs(savePath)
	if err != nil {
		fmt.Printf("SaveAs err:%v\n", err)
		return err
	}
	return nil
}

// 设置第一行
func setExcelFirstRow(f *excelize.File) (err error) {
	date := time.Now().Format("1.2")

	// 合并单元格
	err = f.MergeCell(globalSheetName, "A1", "E1")
	if err != nil {
		return
	}
	// 设置文本
	err = f.SetCellStr(globalSheetName, "A1", "泰豪VR研究院 / VR研发中心每日工作（"+date+"）")
	if err != nil {
		return
	}
	style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size: 16,
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	err = f.SetCellStyle(globalSheetName, "A1", "A1", style)
	if err != nil {
		return
	}
	return nil
}

// 设置excel表头
func setExcelHeader(f *excelize.File) (err error) {
	// 设置文本
	f.SetCellStr(globalSheetName, "A2", "项目名称")
	f.SetCellStr(globalSheetName, "B2", "姓名")
	f.SetCellStr(globalSheetName, "C2", "工作内容")
	f.SetCellStr(globalSheetName, "D2", "昨日进度")
	f.SetCellStr(globalSheetName, "E2", "今日进度")

	style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   14,
			Bold:   true,
			Color:  "#FFFFFF",
			Family: "宋体",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#0070C0"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	err = f.SetCellStyle(globalSheetName, "A2", "E2", style)
	if err != nil {
		return
	}
	return nil
}

// 设置最后一行
func setExcelEndRow(f *excelize.File, lineNum int) (err error) {
	if err = f.SetRowHeight(globalSheetName, lineNum, 35); err != nil {
		return
	}

	// 设置富文本
	if err = f.SetCellRichText(globalSheetName, "A"+strconv.Itoa(lineNum), []excelize.RichTextRun{
		{
			Text: "注：1、如果工作任务无法按时完成请及时向上级反馈",
			Font: &excelize.Font{
				Size:  12,
				Color: "#000000",
			},
		},
		{
			Text: "\r\n    2、昨日进度-NAN = 新增工作任务，今日进度-NAN = 暂无工作安排",
			Font: &excelize.Font{
				Size:  12,
				Color: "#000000",
			},
		},
	}); err != nil {
		return
	}
	style, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText:   true,
			Horizontal: "left",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return
	}
	if err = f.SetCellStyle(globalSheetName, "A"+strconv.Itoa(lineNum), "A"+strconv.Itoa(lineNum), style); err != nil {
		return
	}

	// 合并单元格
	err = f.MergeCell(globalSheetName, "A"+strconv.Itoa(lineNum), "E"+strconv.Itoa(lineNum))
	if err != nil {
		return
	}
	return nil
}

// 设置表格内容
func setTableData(f *excelize.File, projectList []config.ProjectList) (baseIndex int, err error) {
	baseIndex = 3 // 从第三行开始绘制
	for i, project := range projectList {
		var style int
		fillColor := "#BDD7EE"
		if i%2 == 0 {
			fillColor = "#DDEBF7"
		}
		// 绘制项目名
		projectStartIndex := baseIndex
		projectEndIndex := projectStartIndex + project.Total - 1
		for line := projectStartIndex; line <= projectEndIndex; line++ {
			f.SetRowHeight(globalSheetName, line, 22)
		}
		projectStart := "A" + strconv.Itoa(projectStartIndex)
		projectEnd := "A" + strconv.Itoa(projectEndIndex)
		f.MergeCell(globalSheetName, projectStart, projectEnd)
		style, err = f.NewStyle(&excelize.Style{
			Font: &excelize.Font{
				Size: 12,
				Bold: true,
			},
			Alignment: &excelize.Alignment{
				Horizontal: "center",
				Vertical:   "center",
			},
			Fill: excelize.Fill{
				Type:    "pattern",
				Color:   []string{fillColor},
				Pattern: 1,
			},
			Border: []excelize.Border{
				{Type: "top", Color: "000000", Style: 1},
				{Type: "right", Color: "000000", Style: 1},
			},
		})
		err = f.SetCellStyle(globalSheetName, projectStart, projectEnd, style)
		if err != nil {
			return
		}

		// 设置文本
		err = f.SetCellStr(globalSheetName, projectStart, project.ProjectName)
		if err != nil {
			return
		}

		// 绘制姓名
		workIndex := baseIndex
		for _, user := range project.UserList {
			workCount := len(user.WorkList)
			userStartIndex := workIndex
			userEndIndex := userStartIndex + workCount - 1
			userStart := "B" + strconv.Itoa(userStartIndex)
			userEnd := "B" + strconv.Itoa(userEndIndex)
			if workCount > 1 {
				f.MergeCell(globalSheetName, userStart, userEnd)
			}

			style, err = f.NewStyle(&excelize.Style{
				Font: &excelize.Font{
					Size: 12,
				},
				Alignment: &excelize.Alignment{
					Horizontal: "center",
					Vertical:   "center",
				},
				Fill: excelize.Fill{
					Type:    "pattern",
					Color:   []string{fillColor},
					Pattern: 1,
				},
				Border: []excelize.Border{
					{Type: "top", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
				},
			})
			err = f.SetCellStyle(globalSheetName, userStart, userEnd, style)
			if err != nil {
				return
			}

			// 设置文本
			err = f.SetCellStr(globalSheetName, userStart, user.Name)
			if err != nil {
				return
			}

			// 绘制工作内容
			for _, work := range user.WorkList {
				// 绘制标题
				style, err = f.NewStyle(&excelize.Style{
					Font: &excelize.Font{
						Size: 12,
					},
					Alignment: &excelize.Alignment{
						Horizontal: "center",
						Vertical:   "center",
					},
					Fill: excelize.Fill{
						Type:    "pattern",
						Color:   []string{fillColor},
						Pattern: 1,
					},
					Border: []excelize.Border{
						{Type: "top", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
					},
				})
				err = f.SetCellStyle(globalSheetName, "C"+strconv.Itoa(workIndex), "C"+strconv.Itoa(workIndex), style)
				if err != nil {
					return
				}

				// 设置文本
				err = f.SetCellStr(globalSheetName, "C"+strconv.Itoa(workIndex), work.Title)
				if err != nil {
					return
				}

				// 绘制昨日进度
				err = f.SetCellStyle(globalSheetName, "D"+strconv.Itoa(workIndex), "D"+strconv.Itoa(workIndex), style)
				if err != nil {
					return
				}

				// 设置文本
				err = f.SetCellStr(globalSheetName, "D"+strconv.Itoa(workIndex), work.YesterdayProgress)
				if err != nil {
					return
				}

				// 绘制今日进度
				err = f.SetCellStyle(globalSheetName, "E"+strconv.Itoa(workIndex), "E"+strconv.Itoa(workIndex), style)
				if err != nil {
					return
				}

				// 设置文本
				err = f.SetCellStr(globalSheetName, "E"+strconv.Itoa(workIndex), work.TodayProgress)
				if err != nil {
					return
				}

				workIndex += 1
			}
		}

		baseIndex += project.Total
	}
	return
}
