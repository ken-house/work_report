package controller

import (
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"strings"
)

type FileController struct {
}

// 设置全局变量存储目录下的文件列表
var filePathArr = make([]string, 0, 20)

// ScanDirPath 遍历目录下的文件
// @author xudt
func (f *FileController) ScanDirPath(dirPath string) (err error) {
	rd, err := ioutil.ReadDir(dirPath)
	if err != nil {
		fmt.Printf("ioutil.ReadDir err:%v\n", err)
		return err
	}

	for _, file := range rd {
		if file.IsDir() {
			f.ScanDirPath(dirPath + "\\" + file.Name() + "\\")
		} else {
			filePath := dirPath + "\\" + file.Name()
			ext := filepath.Ext(filePath)
			if ext == ".xls" || ext == ".xlsx" {
				filePathArr = append(filePathArr, filePath)
			}
		}
	}

	return nil
}

// getProjectNameFromFileName 从文件路径获取项目名
func getProjectNameFromFilePath(filePath string) string {
	fileName := filepath.Base(filePath)
	position := strings.Index(fileName, "_")
	if position == -1 {
		ext := path.Ext(filePath)
		return strings.ReplaceAll(fileName, ext, "")
	}
	return fileName[0:position]
}

// 判断文件是否存在
func existFile(filePath string) bool {
	_, err := os.Stat(filePath)
	if err != nil {
		if os.IsExist(err) {
			return true
		}
		if os.IsNotExist(err) {
			return false
		}
		return false
	}
	return true
}
