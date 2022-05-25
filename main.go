package main

import (
	"flag"
	"fmt"
	"tapd/work_report/xlsx/config"
	"tapd/work_report/xlsx/controller"
)

var (
	dirPath  string // 要解析的文件目录
	savePath string // 文件存储位置
)

func main() {
	// 接收两个参数
	flag.StringVar(&dirPath, "dir_path", config.DefaultDirPath, "请输入解析文件路径")
	flag.StringVar(&savePath, "save_path", config.DefaultSavePath, "请输入文件存储位置")
	flag.Parse()
	//fmt.Printf("dir_path:%v,save_path:%v\n", dirPath, savePath)

	defer func() {
		if err := recover(); err != nil {
			fmt.Printf("err:%v\n", err)
		}
	}()

	// 1.解析file/input_file/目录下的xlsx文件，生成数据
	var xlsxController = new(controller.XlsxController)
	xlsxData, err := xlsxController.ParseXlsxOrXls(dirPath)
	if err != nil {
		panic(err)
	}
	//fmt.Println(xlsxData)

	// 2.对数据进行format，以便进行下一步操作
	var dataController = new(controller.DataController)
	projectList, err := dataController.Format(xlsxData)
	if err != nil {
		panic(err)
	}
	//fmt.Printf("projectList:%v\n", projectList)

	// 3.生成工作报告xlsx文件
	err = xlsxController.ExportXlsx(savePath, projectList)
	if err != nil {
		panic(err)
	}
	fmt.Printf("xlsx文件生成成功，文件路径：%s\n", savePath)
}
