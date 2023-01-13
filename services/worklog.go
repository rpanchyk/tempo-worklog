package services

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"sort"
	"strings"
	"tempo-worklog/models"
	"time"
)

type WorklogService struct {
	jiraUser                   string
	jiraToken                  string
	tempoToken                 string
	tempoWorklogUrlTemplate    string
	jiraSearchIssueUrlTemplate string
	projectConfigService       *ProjectConfigService
}

func NewWorklogService(jiraUrl, jiraUser, jiraToken, tempoToken string, projectConfigService *ProjectConfigService) *WorklogService {

	return &WorklogService{
		jiraUser:   jiraUser,
		jiraToken:  jiraToken,
		tempoToken: tempoToken,

		// https://api.tempo.io/core/3/worklogs?projectId=PRJ&limit=10&from=2019-12-27&to=2020-07-20
		tempoWorklogUrlTemplate: "https://api.tempo.io/core/3/worklogs?project=%s&from=%s&to=%s&offset=%d&limit=%d",

		// https://company.atlassian.net/rest/api/2/search?fields=summary&jql=key%20in%20(PRJ-384,PRJ-502)&startAt=1&maxResults=1
		jiraSearchIssueUrlTemplate: strings.TrimRight(jiraUrl, "/") + "/rest/api/2/search?fields=summary&jql=key%%20in%%20(%s)&startAt=%d&maxResults=%d",

		projectConfigService: projectConfigService,
	}
}

func (s *WorklogService) GetWorklog(projectKeys []string, dateFrom, dateTo string) (*models.Worklog, error) {
	projectConfigWrapper, err := s.projectConfigService.Get()
	if err != nil {
		return nil, err
	}

	var projects []models.Project

	log.Println("Getting tempo report started")

	for _, projectKey := range projectKeys {
		project, err := s.getProject(projectKey, dateFrom, dateTo, projectConfigWrapper)
		if err != nil {
			return nil, err
		}
		projects = append(projects, *project)
	}

	log.Println("Getting tempo report finished")

	worklog := &models.Worklog{Projects: projects}
	//fmt.Println("worklog", worklog)

	err = s.projectConfigService.Save(projectConfigWrapper, worklog)
	if err != nil {
		return nil, err
	}

	return worklog, nil
}

func (s *WorklogService) getProject(projectKey, dateFrom, dateTo string, projectConfigWrapper *models.ProjectConfigWrapper) (*models.Project, error) {
	// get tempo worklog
	var tempoResults []models.TempoResult
	offset := 0
	limit := 100

	for {
		response, err := s.getTempoWorklog(projectKey, dateFrom, dateTo, offset, limit)
		if err != nil {
			return nil, err
		}

		//fmt.Println("Tempo response:", response)
		log.Println("Fetched tempo report for", projectKey, "project:", response.Metadata.Count, "records")

		tempoResults = append(tempoResults, response.Results...)

		if response.Metadata.Count < response.Metadata.Limit {
			break
		}
		offset = response.Metadata.Count + response.Metadata.Offset
	}
	//fmt.Println("tempoResults", tempoResults)

	// convert tempo worklog to internal structure
	projectConfig := projectConfigWrapper.ProjectKeyToConfig[projectKey]
	users, err := s.getUsers(tempoResults, &projectConfig)
	if err != nil {
		return nil, err
	}

	return &models.Project{Key: projectKey, Users: users}, nil
}

func (s *WorklogService) getTempoWorklog(projectKey, dateFrom, dateTo string, offset, limit int) (*models.TempoResponse, error) {
	url := fmt.Sprintf(s.tempoWorklogUrlTemplate, projectKey, dateFrom, dateTo, offset, limit)

	client := http.Client{Timeout: time.Second * 60}

	request, err := http.NewRequest(http.MethodGet, url, nil)
	if err != nil {
		return nil, err
	}

	request.Header.Set("Authorization", "Bearer "+s.tempoToken)
	request.Header.Set("Content-Type", "application/json")

	response, err := client.Do(request)
	if err != nil {
		return nil, err
	}

	if response.Body != nil {
		defer response.Body.Close()
	}

	body, err := ioutil.ReadAll(response.Body)
	if err != nil {
		return nil, err
	}

	tempoResponse := &models.TempoResponse{}
	err = json.Unmarshal(body, tempoResponse)
	if err != nil {
		return nil, err
	}

	return tempoResponse, nil
}

func (s *WorklogService) getUsers(results []models.TempoResult, projectConfig *models.ProjectConfig) ([]models.User, error) {
	userIdToTempoResult := map[string][]models.TempoResult{} // group tempo results by account id

	for _, result := range results {
		accountId := result.Author.AccountId
		if _, ok := userIdToTempoResult[accountId]; ok {
			userIdToTempoResult[accountId] = append(userIdToTempoResult[accountId], result)
		} else {
			userIdToTempoResult[accountId] = []models.TempoResult{result}
		}
	}
	//fmt.Println("userIdToTempoResult", userIdToTempoResult)

	var users []models.User

	for _, userResults := range userIdToTempoResult {
		issues, err := s.getIssues(userResults)
		if err != nil {
			return nil, err
		}

		author := userResults[0].Author
		userConfig := projectConfig.UserNameToConfig[author.DisplayName]

		user := models.User{
			AccountId:   author.AccountId,
			DisplayName: author.DisplayName,
			Position:    userConfig.Position,
			Rate:        userConfig.Rate,
			Issues:      issues,
		}
		users = append(users, user)
	}

	sort.Slice(users, func(i, j int) bool {
		return strings.ToLower(users[i].DisplayName) < strings.ToLower(users[j].DisplayName)
	})
	//fmt.Println("users", users)

	return users, nil
}

func (s *WorklogService) getIssues(results []models.TempoResult) ([]models.Issue, error) {
	issueKeyToResults := map[string][]models.TempoResult{} // group by issue id

	for _, result := range results {
		issueKey := result.Issue.Key

		if _, ok := issueKeyToResults[issueKey]; ok {
			issueKeyToResults[issueKey] = append(issueKeyToResults[issueKey], result)
		} else {
			issueKeyToResults[issueKey] = []models.TempoResult{result}
		}
	}

	issueKeyToSummary, err := s.getIssueKeyToSummary(issueKeyToResults)
	if err != nil {
		return nil, err
	}

	var issues []models.Issue

	for _, results := range issueKeyToResults {
		issueKey := results[0].Issue.Key

		efforts, err := s.getEfforts(results)
		if err != nil {
			return nil, err
		}

		issue := models.Issue{
			Key:     issueKey,
			Summary: issueKeyToSummary[issueKey],
			Efforts: efforts,
		}

		issues = append(issues, issue)
	}
	//fmt.Println("issues", issues)

	return issues, nil
}

func (s *WorklogService) getIssueKeyToSummary(issueKeyToResults map[string][]models.TempoResult) (map[string]string, error) {
	separator := ""
	keys := ""
	for key := range issueKeyToResults {
		if keys != "" {
			separator = ","
		}
		keys += separator + key
	}

	issueKeyToSummary := map[string]string{}
	offset := 0
	limit := 100

	for {
		response, err := s.searchIssueByKeys(keys, offset, limit)
		if err != nil {
			return nil, err
		}
		//fmt.Println("searchIssueByKeys", response)

		for _, issue := range response.Issues {
			issueKeyToSummary[issue.Key] = issue.Fields.Summary
		}

		count := response.StartAt + len(response.Issues)
		if count >= response.Total {
			break
		}
		offset = count
	}
	//fmt.Println("issueKeyToSummary", issueKeyToSummary)

	return issueKeyToSummary, nil
}

func (s *WorklogService) searchIssueByKeys(keys string, offset, limit int) (*models.JiraSearchIssueResponse, error) {
	url := fmt.Sprintf(s.jiraSearchIssueUrlTemplate, keys, offset, limit)

	client := http.Client{Timeout: time.Second * 60}

	request, err := http.NewRequest(http.MethodGet, url, nil)
	if err != nil {
		return nil, err
	}

	encodedToken := base64.URLEncoding.EncodeToString([]byte(s.jiraUser + ":" + s.jiraToken))
	request.Header.Set("Authorization", "Basic "+encodedToken)
	request.Header.Set("Content-Type", "application/json")

	response, err := client.Do(request)
	if err != nil {
		return nil, err
	}

	if response.Body != nil {
		defer response.Body.Close()
	}

	body, err := ioutil.ReadAll(response.Body)
	if err != nil {
		return nil, err
	}

	jiraSearchIssueResponse := &models.JiraSearchIssueResponse{}
	err = json.Unmarshal(body, jiraSearchIssueResponse)
	if err != nil {
		return nil, err
	}

	return jiraSearchIssueResponse, nil
}

func (s *WorklogService) getEfforts(results []models.TempoResult) ([]models.Effort, error) {
	dateToEffort := map[string]models.Effort{}

	for _, result := range results {
		date := result.StartDate

		var timeSpentSeconds int
		if _, ok := dateToEffort[date]; ok {
			prevEffort := dateToEffort[date]
			timeSpentSeconds = prevEffort.TimeSpentSeconds + result.TimeSpentSeconds
		} else {
			timeSpentSeconds = result.TimeSpentSeconds
		}

		dateToEffort[date] = models.Effort{
			Date:             result.StartDate,
			TimeSpentSeconds: timeSpentSeconds,
		}
	}

	var efforts []models.Effort
	for _, effort := range dateToEffort {
		efforts = append(efforts, effort)
	}

	return efforts, nil
}
