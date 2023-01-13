package models

// worklog -> projects -> users -> issues -> efforts

type Worklog struct {
	Projects []Project
}

type Project struct {
	Key   string
	Users []User
}

type User struct {
	AccountId   string
	DisplayName string
	Position    string
	Rate        int
	Issues      []Issue
}

type Issue struct {
	Key     string
	Summary string
	Efforts []Effort
}

type Effort struct {
	Date             string
	TimeSpentSeconds int
}
