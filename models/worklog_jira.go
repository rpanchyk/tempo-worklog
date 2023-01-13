package models

type JiraSearchIssueResponse struct {
	Total      int               `json:"total"`
	StartAt    int               `json:"startAt"`
	MaxResults int               `json:"maxResults"`
	Issues     []JiraSearchIssue `json:"issues"`
}

type JiraSearchIssue struct {
	Key    string                `json:"key"`
	Fields JiraSearchIssueFields `json:"fields"`
}

type JiraSearchIssueFields struct {
	Summary string `json:"summary"`
}
