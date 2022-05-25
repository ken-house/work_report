package config

// XlsxData 解析项目文件后的数据格式
type XlsxData struct {
	Project   string  // 项目名
	StoryList []Story // 需求列表
}

// Story 需求数据
type Story struct {
	Title string
	Owner string
	Begin string
	Due   string
}

// ProjectList 导出文件项目列表结构体
type ProjectList struct {
	ProjectName string
	Total       int // 总条数
	UserList    []UserList
}

// UserList 导出文件用户列表结构体
type UserList struct {
	Name     string
	WorkList []WorkList
}

// WorkList 导出文件需求列表结构体
type WorkList struct {
	Title             string
	YesterdayProgress string
	TodayProgress     string
}
