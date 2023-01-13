package services

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"sort"
	"strconv"
	"tempo-worklog/models"
	"tempo-worklog/utils"
)

type ProjectConfigService struct {
	filePath string
}

func NewProjectConfigService(filePath string) *ProjectConfigService {
	return &ProjectConfigService{
		filePath: filePath,
	}
}

func (s *ProjectConfigService) Get() (*models.ProjectConfigWrapper, error) {
	_, err := os.Stat(s.filePath)
	if err != nil && errors.Is(err, os.ErrNotExist) {
		return &models.ProjectConfigWrapper{}, nil
	}

	f, err := excelize.OpenFile(s.filePath)
	if err != nil {
		return nil, err
	}

	projectKeyToConfig := map[string]models.ProjectConfig{}

	for _, sheet := range f.GetSheetList() {
		userConfigs, err := s.getUserConfigs(f, sheet)
		if err != nil {
			return nil, err
		}

		projectKeyToConfig[sheet] = models.ProjectConfig{UserNameToConfig: userConfigs}
	}

	err = f.Close()
	if err != nil {
		return nil, err
	}

	projectConfigWrapper := &models.ProjectConfigWrapper{ProjectKeyToConfig: projectKeyToConfig}

	log.Println("Parsed", s.filePath, utils.ToPrettyString("config", projectConfigWrapper))

	return projectConfigWrapper, nil
}

func (s *ProjectConfigService) getUserConfigs(f *excelize.File, sheet string) (map[string]models.UserConfig, error) {
	userConfigs := map[string]models.UserConfig{}

	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}

	rowIndex := 0
	for rows.Next() {
		rowIndex++
		if rowIndex <= 1 { // skip header rows
			continue
		}

		rowCols, err := rows.Columns()
		if err != nil {
			return nil, err
		}

		if rowCols == nil {
			continue
		}

		userName := ""
		if len(rowCols) > 0 {
			userName = rowCols[0]
		}
		if len(userName) == 0 { // required
			continue
		}

		position := ""
		if len(rowCols) > 1 {
			position = rowCols[1]
		}

		rate := 0
		if len(rowCols) > 2 {
			rate, err = strconv.Atoi(rowCols[2])
			if err != nil {
				return nil, err
			}
		}

		userConfigs[userName] = models.UserConfig{
			Position: position,
			Rate:     rate,
		}
	}

	if err = rows.Close(); err != nil {
		return nil, err
	}

	return userConfigs, nil
}

func (s *ProjectConfigService) Save(projectConfigWrapper *models.ProjectConfigWrapper, worklog *models.Worklog) error {
	updatedProjectConfigWrapper := s.updateProjectConfigs(projectConfigWrapper, worklog)

	err := s.saveProjectConfigs(updatedProjectConfigWrapper)
	if err != nil {
		return err
	}

	log.Println("Synchronized", s.filePath, utils.ToPrettyString("config", projectConfigWrapper))

	return nil
}

func (s *ProjectConfigService) updateProjectConfigs(projectConfigWrapper *models.ProjectConfigWrapper, worklog *models.Worklog) *models.ProjectConfigWrapper {
	updatedProjectConfigs := map[string]models.ProjectConfig{}

	for _, worklogProject := range worklog.Projects {
		userConfigs := projectConfigWrapper.ProjectKeyToConfig[worklogProject.Key].UserNameToConfig
		updatedUserConfigs := map[string]models.UserConfig{}

		for _, worklogUser := range worklogProject.Users {
			updatedUserConfig := models.UserConfig{}

			if userConfig, ok := userConfigs[worklogUser.DisplayName]; ok {
				updatedUserConfig = models.UserConfig{Position: userConfig.Position, Rate: userConfig.Rate}
			}
			updatedUserConfigs[worklogUser.DisplayName] = updatedUserConfig
		}

		updatedProjectConfigs[worklogProject.Key] = models.ProjectConfig{UserNameToConfig: updatedUserConfigs}
	}

	return &models.ProjectConfigWrapper{ProjectKeyToConfig: updatedProjectConfigs}
}

func (s *ProjectConfigService) saveProjectConfigs(projectConfigWrapper *models.ProjectConfigWrapper) error {
	f := excelize.NewFile()

	projectConfigs := projectConfigWrapper.ProjectKeyToConfig

	projectKeys := make([]string, 0, len(projectConfigs))
	for projectKey := range projectConfigs {
		projectKeys = append(projectKeys, projectKey)
	}
	sort.Strings(projectKeys)

	for _, project := range projectKeys {
		err := s.createProjectSheet(f, project)
		if err != nil {
			return err
		}

		err = s.fillProjectSheetWithUsers(f, project, projectConfigs[project].UserNameToConfig)
		if err != nil {
			return err
		}
	}

	err := f.SaveAs(s.filePath)
	if err != nil {
		return err
	}

	err = f.Close()
	if err != nil {
		return err
	}

	return nil
}

func (s *ProjectConfigService) createProjectSheet(f *excelize.File, sheet string) error {
	sheetIndex, err := f.GetSheetIndex("Sheet1")
	if err != nil {
		return err
	}
	if sheetIndex != -1 {
		f.SetSheetName("Sheet1", sheet)
	} else {
		f.NewSheet(sheet)
	}

	font := excelize.Font{Size: 12}
	style, err := f.NewStyle(&excelize.Style{Font: &font})
	err = f.SetColStyle(sheet, "A", style)
	if err != nil {
		return err
	}

	style, err = f.NewStyle(&excelize.Style{Font: &font})
	err = f.SetColStyle(sheet, "B", style)
	if err != nil {
		return err
	}

	style, err = f.NewStyle(&excelize.Style{Font: &font, NumFmt: 177})
	err = f.SetColStyle(sheet, "C", style)
	if err != nil {
		return err
	}

	alignment := excelize.Alignment{Horizontal: "center", Vertical: "center"}
	borders := []excelize.Border{
		{Type: "top", Color: "#000000", Style: 1},
		{Type: "left", Color: "#000000", Style: 1},
		{Type: "bottom", Color: "#000000", Style: 1},
		{Type: "right", Color: "#000000", Style: 1},
	}
	headerFont := excelize.Font{Size: 13, Color: "#ffffff", Bold: true}
	fill := excelize.Fill{Color: []string{"#009a00"}, Type: "pattern", Pattern: 3}
	style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &headerFont, Border: borders, Fill: fill})
	err = f.SetCellStyle(sheet, "A1", "C1", style)
	if err != nil {
		return err
	}

	err = f.SetRowHeight(sheet, 1, 25)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, "A", "A", 30)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, "B", "B", 30)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, "C", "C", 10)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "A1", "Name")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "B1", "Position")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "C1", "Rate")
	if err != nil {
		return err
	}

	return nil
}

func (s *ProjectConfigService) fillProjectSheetWithUsers(f *excelize.File, sheet string, users map[string]models.UserConfig) error {
	rowsCount, err := s.getRowsCountInColumn(f, sheet, 1)
	if err != nil {
		return err
	}

	lastRowIndex := *rowsCount

	userNames := make([]string, 0, len(users))
	for userName := range users {
		userNames = append(userNames, userName)
	}
	sort.Strings(userNames)

	conditionalFormat, err := s.getConditionalFormat(f)
	if err != nil {
		return err
	}

	for _, userName := range userNames {
		lastRowIndex++
		rowIndex := strconv.Itoa(lastRowIndex)

		err = f.SetConditionalFormat(sheet, "C"+rowIndex, conditionalFormat)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, "A"+rowIndex, userName)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, "B"+rowIndex, users[userName].Position)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, "C"+rowIndex, users[userName].Rate)
		if err != nil {
			return err
		}
	}

	return nil
}

func (s *ProjectConfigService) getRowsCountInColumn(f *excelize.File, sheet string, col int) (*int, error) {
	cols, err := f.Cols(sheet)
	if err != nil {
		return nil, err
	}

	// shift to column
	for i := 0; i < col; i++ {
		cols.Next()
	}

	colRows, err := cols.Rows()
	if err != nil {
		return nil, err
	}

	result := len(colRows)
	return &result, nil
}

func (s *ProjectConfigService) getConditionalFormat(f *excelize.File) ([]excelize.ConditionalFormatOptions, error) {
	format, err := f.NewConditionalStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#9A0511"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#FEC7CE"}, Pattern: 1},
	})
	if err != nil {
		return nil, err
	}

	return []excelize.ConditionalFormatOptions{{Type: "cell", Criteria: "=", Format: format, Value: "0"}}, nil
}
