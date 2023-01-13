package models

type TempoResponse struct {
	Metadata TempoMetadata `json:"metadata"`
	Results  []TempoResult `json:"results"`
}

type TempoMetadata struct {
	Count  int `json:"count"`
	Offset int `json:"offset"`
	Limit  int `json:"limit"`
}

type TempoResult struct {
	Author           TempoAuthor `json:"author"`
	Issue            TempoIssue  `json:"issue"`
	StartDate        string      `json:"startDate"`
	TimeSpentSeconds int         `json:"timeSpentSeconds"`
}

type TempoAuthor struct {
	AccountId   string `json:"accountId"`
	DisplayName string `json:"displayName"`
}

type TempoIssue struct {
	Key string `json:"key"`
}
