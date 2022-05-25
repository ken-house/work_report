package controller

import (
	"errors"
	"math"
	"strconv"
	"strings"
	"tapd/work_report/xlsx/config"
	"tapd/work_report/xlsx/tools"
	"time"
)

type DataController struct {
}

// Format 格式化数据
// @author ken
func (d *DataController) Format(xlsxData []config.XlsxData) (projectList []config.ProjectList, err error) {
	projectList = make([]config.ProjectList, 0, 30)
	for _, v := range xlsxData {
		workCount := 0
		userMap := make(map[string][]config.WorkList, 100)
		for _, story := range v.StoryList {
			ownerList := strings.Split(story.Owner, ";")
			for _, owner := range ownerList {
				workCount++
				work := config.WorkList{
					Title:             story.Title,
					YesterdayProgress: calProgress(story.Begin, story.Due, time.Now().AddDate(0, 0, -1).Unix()),
					TodayProgress:     calProgress(story.Begin, story.Due, time.Now().Unix()),
				}
				if workList, ok := userMap[owner]; !ok { // 不存在该用户则置空
					workList := make([]config.WorkList, 0, 20)
					workList = append(workList, work)
					userMap[owner] = workList
				} else { // 存在则取出值，并追加一个story
					workList = append(workList, work)
					userMap[owner] = workList
				}
			}
		}

		// userMap转为slice
		userList := make([]config.UserList, 0, 100)
		for userName, workList := range userMap {
			user := config.UserList{
				Name:     userName,
				WorkList: workList,
			}
			userList = append(userList, user)
		}

		project := config.ProjectList{
			ProjectName: "《" + v.Project + "》",
			Total:       workCount,
			UserList:    userList,
		}
		projectList = append(projectList, project)
	}

	return projectList, nil
}

func (d *DataController) FormatDataFromXlsx(filePath string, rows [][]string) (data config.XlsxData, err error) {
	// 所需解析文件的表头
	xlsxHeader := []string{"标题", "处理人", "预计开始", "预计结束", "父需求"}
	headerIndex := make([]int, 0, 10)
	if len(rows) == 0 {
		return data, errors.New("解析数据为空")
	}
	headerSize := len(rows[0]) // 通过表头读出共多少列
	// 检查是否包含这五列
	for _, header := range xlsxHeader {
		if !tools.IsContain(rows[0], header) {
			return data, errors.New("文件数据格式错误")
		}
	}
	for _, header := range xlsxHeader {
		for i, title := range rows[0] {
			if title == header {
				headerIndex = append(headerIndex, i)
			}
		}
	}

	// 根据对应的索引位置，找出对应的值组成正确的数据格式
	invalidStoryList := make([][]string, 0, 100)
	for i, row := range rows {
		if i == 0 {
			continue
		}
		// 保证slice长度一致
		if len(row) < headerSize {
			for m := len(row); m < headerSize; m++ {
				row = append(row, "")
			}
		}

		invalidRow := make([]string, 0, 30)
		for _, index := range headerIndex {
			invalidRow = append(invalidRow, row[index])
		}
		invalidStoryList = append(invalidStoryList, invalidRow)
	}

	storyList := make([]config.Story, 0, 100)
	for _, row := range invalidStoryList {
		if row[4] != "" { // 子需求跳过
			continue
		}
		story := config.Story{
			Title: row[0],
			Owner: row[1],
			Begin: row[2],
			Due:   row[3],
		}
		storyList = append(storyList, story)
	}
	data = config.XlsxData{
		Project:   getProjectNameFromFilePath(filePath),
		StoryList: storyList,
	}
	return data, nil
}

// calProgress 计算进度
// @author ken
func calProgress(beginDate string, endDate string, timestamp int64) string {
	// 1.需求开始和结束日期的天数 - 总工期
	if beginDate == "" || endDate == "" {
		return "NAN"
	}

	loc, _ := time.LoadLocation("Local") //获取时区
	bd, _ := time.ParseInLocation("2006-01-02", beginDate, loc)
	beginTime := bd.Unix()
	ed, _ := time.ParseInLocation("2006-01-02", endDate, loc)
	endTime := ed.Unix()
	totalDay := (endTime-beginTime)/86400 + 1

	// 2.指定时间戳距离需求开始日期的天数 - 已使用工期
	// 今天 > 结束日期
	if timestamp >= endTime {
		return "100%"
	}
	// 今天 < 开始日期
	if timestamp < beginTime {
		return "NAN"
	}
	dueDay := (timestamp-beginTime)/86400 + 1

	// 3.计算已使用工期所占百分比，保留整数
	progress := math.Ceil(float64(dueDay) / float64(totalDay) * 100)

	return strconv.Itoa(int(progress)) + "%"
}
