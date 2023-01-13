package services

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"math"
	"strconv"
	"strings"
	"tempo-worklog/constants"
	"tempo-worklog/models"
	"time"
)

const (
	ColumnReportDateFormat = "%d/%d"
)

type ExcelService struct {
	filePath string
}

func NewExcelService(filePath string) *ExcelService {
	return &ExcelService{filePath: filePath}
}

func (s *ExcelService) Save(worklog *models.Worklog, dateFrom, dateTo string) error {
	f := excelize.NewFile()

	sheet, err := s.getSheetName(dateFrom, dateTo)
	if err != nil {
		return err
	}

	// prepare
	err = s.prepare(f, *sheet, dateFrom, dateTo)
	if err != nil {
		return err
	}

	// fill data
	err = s.fill(f, *sheet, dateFrom, dateTo, worklog)
	if err != nil {
		return err
	}

	// save to file
	err = f.SaveAs(s.filePath)
	if err != nil {
		return err
	}

	err = f.Close()
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) getSheetName(dateFrom, dateTo string) (*string, error) {
	startDate, err := time.Parse(constants.InputDateFormat, dateFrom)
	if err != nil {
		return nil, err
	}

	endDate, err := time.Parse(constants.InputDateFormat, dateTo)
	if err != nil {
		return nil, err
	}

	sheetName := startDate.Format("Jan 2") + " - " + endDate.Format("Jan 2")
	return &sheetName, nil
}

func (s *ExcelService) prepare(f *excelize.File, sheet, dateFrom, dateTo string) error {
	f.SetSheetName("Sheet1", sheet)

	alignment := excelize.Alignment{Horizontal: "center", Vertical: "center"}
	font := excelize.Font{Size: 13, Color: "#000000", Bold: true}
	style1, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font})
	err = f.SetRowStyle(sheet, 1, 1, style1)
	if err != nil {
		return err
	}

	headerBorders := []excelize.Border{
		{Type: "top", Color: "#000000", Style: 1},
		{Type: "left", Color: "#000000", Style: 1},
		{Type: "bottom", Color: "#000000", Style: 1},
		{Type: "right", Color: "#000000", Style: 1},
	}
	headerFont := excelize.Font{Size: 13, Color: "#ffffff", Bold: true}
	headerFill := excelize.Fill{Color: []string{"#2487bc"}, Type: "pattern", Pattern: 3}
	style, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &headerFont, Border: headerBorders, Fill: headerFill})
	err = f.SetCellStyle(sheet, "A1", "F1", style)
	if err != nil {
		return err
	}

	err = f.SetRowHeight(sheet, 1, 25)
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

	err = f.SetCellValue(sheet, "C1", "Task")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "D1", "Rate")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "E1", "Hours")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "F1", "Total cost")
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

	err = f.SetColWidth(sheet, "C", "C", 60)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, "D", "D", 10)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, "E", "E", 10)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, "F", "F", 20)
	if err != nil {
		return err
	}

	// https://xuri.me/excelize/en/utils.html#SetPanes
	err = f.SetPanes(sheet, &excelize.Panes{Freeze: true, XSplit: 6, YSplit: 1})
	if err != nil {
		return err
	}

	// dates
	startDate, err := time.Parse(constants.InputDateFormat, dateFrom)
	if err != nil {
		return err
	}

	endDate, err := time.Parse(constants.InputDateFormat, dateTo)
	if err != nil {
		return err
	}

	i := 6
	for date := startDate; !date.After(endDate); date = date.AddDate(0, 0, 1) {
		formattedDate := fmt.Sprintf(ColumnReportDateFormat, date.Month(), date.Day())
		i++

		cell, err := excelize.CoordinatesToCellName(i, 1)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, cell, formattedDate)
		if err != nil {
			return err
		}

		col := strings.TrimRight(cell, "1")

		err = f.SetColWidth(sheet, col, col, 6)
		if err != nil {
			return err
		}

		alignment := excelize.Alignment{Horizontal: "center", Vertical: "center"}
		style, err := f.NewStyle(&excelize.Style{Alignment: &alignment})
		err = f.SetColStyle(sheet, col, style)
		if err != nil {
			return err
		}

		style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Border: headerBorders, Font: &font})
		err = f.SetCellStyle(sheet, cell, cell, style)
		if err != nil {
			return err
		}
	}

	return nil
}

func (s *ExcelService) fill(f *excelize.File, sheet string, dateFrom, dateTo string, worklog *models.Worklog) error {
	colsCount, err := s.getColsCountInRow(f, sheet, 1)
	if err != nil {
		return err
	}
	//fmt.Println("colsCount", *colsCount)

	context := models.ExcelContext{
		FirstDateColumnIndex: 7,
		ColsCount:            *colsCount,
		LastRowIndex:         1,
	}

	for _, project := range worklog.Projects {
		log.Println("Creating report for project:", project.Key)

		err = s.fillProjectRow(f, sheet, project.Key, &context)
		if err != nil {
			return err
		}

		for _, user := range project.Users {
			log.Println("Processing issues for user:", user.DisplayName, fmt.Sprintf("(%d)", len(user.Issues)))

			err = s.fillUserRow(f, sheet, user, &context)
			if err != nil {
				return err
			}

			for _, issue := range user.Issues {
				err = s.fillIssueRow(f, sheet, issue, user.Rate, &context)
				if err != nil {
					return err
				}
			}
		}
	}

	err = s.fillTotalRow(f, sheet, worklog, &context)
	if err != nil {
		return err
	}

	err = s.finalize(f, sheet, dateFrom, dateTo, &context)
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) getColsCountInRow(f *excelize.File, sheet string, row int) (*int, error) {
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}

	if !rows.Next() {
		return nil, errors.New("no rows")
	}

	rowColumns, err := rows.Columns()
	if err != nil {
		return nil, err
	}

	if err = rows.Close(); err != nil {
		return nil, err
	}

	result := len(rowColumns)
	return &result, nil
}

func (s *ExcelService) fillProjectRow(f *excelize.File, sheet string, projectKey string, context *models.ExcelContext) error {
	context.LastRowIndex++

	cell, err := excelize.CoordinatesToCellName(1, context.LastRowIndex)
	if err != nil {
		return err
	}

	cellTo, err := excelize.CoordinatesToCellName(context.ColsCount, context.LastRowIndex)
	if err != nil {
		return err
	}

	err = f.SetRowHeight(sheet, context.LastRowIndex, 25)
	if err != nil {
		return err
	}

	alignment := excelize.Alignment{Vertical: "center"}
	font := excelize.Font{Size: 12, Color: "#000000", Bold: true}
	fill := excelize.Fill{Color: []string{"#bee0f2"}, Type: "pattern", Pattern: 3}
	style, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font, Fill: fill})
	err = f.SetCellStyle(sheet, cell, cellTo, style)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, cell, projectKey)
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) fillUserRow(f *excelize.File, sheet string, user models.User, context *models.ExcelContext) error {
	context.LastRowIndex++

	font := excelize.Font{Size: 12, Color: "#000000", Bold: true}
	fill := excelize.Fill{Color: []string{"#d3e2ea"}, Type: "pattern", Pattern: 3}
	style, err := f.NewStyle(&excelize.Style{Font: &font, Fill: fill})
	err = f.SetRowStyle(sheet, context.LastRowIndex, context.LastRowIndex, style)
	if err != nil {
		return err
	}

	rowIndex := strconv.Itoa(context.LastRowIndex)

	err = f.SetCellValue(sheet, "A"+rowIndex, user.DisplayName)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "B"+rowIndex, user.Position)
	if err != nil {
		return err
	}

	font = excelize.Font{Size: 12, Color: "#000000", Bold: true}
	fill = excelize.Fill{Color: []string{"#d3e2ea"}, Type: "pattern", Pattern: 3}
	style, err = f.NewStyle(&excelize.Style{Font: &font, Fill: fill, NumFmt: 177})
	err = f.SetCellStyle(sheet, "D"+rowIndex, "D"+rowIndex, style)
	if err != nil {
		return err
	}

	conditionalFormat, err := s.getConditionalFormat(f)
	if err != nil {
		return err
	}

	err = f.SetConditionalFormat(sheet, "D"+rowIndex, conditionalFormat)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "D"+rowIndex, user.Rate)
	if err != nil {
		return err
	}

	// hours
	fill = excelize.Fill{Color: []string{"#d3e2ea"}, Type: "pattern", Pattern: 3}
	style, err = f.NewStyle(&excelize.Style{Fill: fill})
	err = f.SetCellStyle(sheet, "E"+rowIndex, "E"+rowIndex, style)
	if err != nil {
		return err
	}

	firstCol, err := excelize.ColumnNumberToName(context.FirstDateColumnIndex)
	if err != nil {
		return err
	}

	lastCol, err := excelize.ColumnNumberToName(context.ColsCount)
	if err != nil {
		return err
	}

	formula := "sum(" + firstCol + rowIndex + ":" + lastCol + rowIndex + ")"

	err = f.SetCellFormula(sheet, "E"+rowIndex, formula)
	if err != nil {
		return err
	}

	// total cost
	fill = excelize.Fill{Color: []string{"#d3e2ea"}, Type: "pattern", Pattern: 3}
	style, err = f.NewStyle(&excelize.Style{Fill: fill, NumFmt: 177})
	err = f.SetCellStyle(sheet, "F"+rowIndex, "F"+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellFormula(sheet, "F"+rowIndex, "D"+rowIndex+"*"+"E"+rowIndex)
	if err != nil {
		return err
	}

	for i := context.FirstDateColumnIndex; i <= context.ColsCount; i++ {
		col, err := excelize.ColumnNumberToName(i)
		if err != nil {
			return err
		}

		alignment := excelize.Alignment{Horizontal: "center"}
		style, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font, Fill: fill})
		err = f.SetCellStyle(sheet, col+rowIndex, col+rowIndex, style)
		if err != nil {
			return err
		}

		columnDate, err := f.GetCellValue(sheet, col+"1")
		if err != nil {
			return err
		}

		isAnyValuePresent := false

		for _, issue := range user.Issues {
			for _, effort := range issue.Efforts {
				date, err := time.Parse(constants.InputDateFormat, effort.Date)
				if err != nil {
					return err
				}

				formattedDate := fmt.Sprintf(ColumnReportDateFormat, date.Month(), date.Day())

				if columnDate == formattedDate {
					isAnyValuePresent = true
					break
				}
			}
		}

		firstRowIndex := context.LastRowIndex + 1
		lastRowIndex := context.LastRowIndex + len(user.Issues)

		if isAnyValuePresent {
			formula := "sum(" + col + strconv.Itoa(firstRowIndex) + ":" + col + strconv.Itoa(lastRowIndex) + ")"

			err = f.SetCellFormula(sheet, col+rowIndex, formula, excelize.FormulaOpts{})
			if err != nil {
				return err
			}
		}
	}

	return nil
}

func (s *ExcelService) getConditionalFormat(f *excelize.File) ([]excelize.ConditionalFormatOptions, error) {
	format, err := f.NewConditionalStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#9A0511"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#FEC7CE"}, Pattern: 1},
	})
	if err != nil {
		return nil, err
	}

	return []excelize.ConditionalFormatOptions{{Type: "cell", Criteria: "=", Format: format, Value: "0"}}, nil
}

func (s *ExcelService) fillIssueRow(f *excelize.File, sheet string, issue models.Issue, rate int, context *models.ExcelContext) error {
	context.LastRowIndex++
	rowIndex := strconv.Itoa(context.LastRowIndex)

	//fmt.Println("issue", issue)

	alignment := excelize.Alignment{WrapText: true}
	style, err := f.NewStyle(&excelize.Style{Alignment: &alignment})
	err = f.SetCellStyle(sheet, "C"+rowIndex, "C"+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "C"+rowIndex, issue.Key+": "+issue.Summary)
	if err != nil {
		return err
	}

	style, err = f.NewStyle(&excelize.Style{NumFmt: 177})
	err = f.SetCellStyle(sheet, "D"+rowIndex, "D"+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "D"+rowIndex, rate)
	if err != nil {
		return err
	}

	// hours
	firstCol, err := excelize.ColumnNumberToName(context.FirstDateColumnIndex)
	if err != nil {
		return err
	}

	lastCol, err := excelize.ColumnNumberToName(context.ColsCount)
	if err != nil {
		return err
	}

	formula := "sum(" + firstCol + rowIndex + ":" + lastCol + rowIndex + ")"

	err = f.SetCellFormula(sheet, "E"+rowIndex, formula)
	if err != nil {
		return err
	}

	// total cost
	style, err = f.NewStyle(&excelize.Style{NumFmt: 177})
	err = f.SetCellStyle(sheet, "F"+rowIndex, "F"+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellFormula(sheet, "F"+rowIndex, "D"+rowIndex+"*"+"E"+rowIndex)
	if err != nil {
		return err
	}

	for _, effort := range issue.Efforts {
		date, err := time.Parse(constants.InputDateFormat, effort.Date)
		if err != nil {
			return err
		}

		formattedDate := fmt.Sprintf(ColumnReportDateFormat, date.Month(), date.Day())

		for i := 7; i <= context.ColsCount; i++ {
			col, err := excelize.ColumnNumberToName(i)
			if err != nil {
				return err
			}

			value, err := f.GetCellValue(sheet, col+"1")
			if err != nil {
				return err
			}

			if value == formattedDate {
				err = f.SetCellValue(sheet, col+rowIndex, s.convertSecondsToHours(effort.TimeSpentSeconds))
				if err != nil {
					return err
				}
				break
			}
		}
	}

	return nil
}

func (s *ExcelService) convertSecondsToHours(seconds int) float64 {
	value := float64(seconds) / 3600

	return math.Round(value*100) / 100
}

func (s *ExcelService) fillTotalRow(f *excelize.File, sheet string, worklog *models.Worklog, context *models.ExcelContext) error {
	context.LastRowIndex++
	rowIndex := strconv.Itoa(context.LastRowIndex)

	alignment := excelize.Alignment{Horizontal: "center", Vertical: "center"}
	font := excelize.Font{Size: 12, Color: "#000000", Bold: true}
	fill := excelize.Fill{Color: []string{"#53aede"}, Type: "pattern", Pattern: 3}
	style, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font, Fill: fill})
	err = f.SetRowStyle(sheet, context.LastRowIndex, context.LastRowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, "A"+rowIndex, "Total")
	if err != nil {
		return err
	}

	totalHoursFormula := ""
	totalCostFormula := ""

	for i := 2; i < context.LastRowIndex; i++ {
		isUserRow, err := s.isUserRow(f, sheet, worklog, i)
		if err != nil {
			return err
		}

		if *isUserRow {
			separator := ""
			if len(totalHoursFormula) > 0 {
				separator = ","
			}

			totalHoursFormula += separator + "E" + strconv.Itoa(i)
			totalCostFormula += separator + "F" + strconv.Itoa(i)
		}
	}

	style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font, Fill: fill})
	err = f.SetCellStyle(sheet, "E"+rowIndex, "E"+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellFormula(sheet, "E"+rowIndex, "sum("+totalHoursFormula+")")
	if err != nil {
		return err
	}

	style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font, Fill: fill, NumFmt: 177})
	err = f.SetCellStyle(sheet, "F"+rowIndex, "F"+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellFormula(sheet, "F"+rowIndex, "sum("+totalCostFormula+")")
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) isUserRow(f *excelize.File, sheet string, worklog *models.Worklog, rowIndex int) (*bool, error) {
	userNameCandidate, err := f.GetCellValue(sheet, "A"+strconv.Itoa(rowIndex))
	if err != nil {
		return nil, err
	}

	falseResult := false
	trueResult := true

	if len(userNameCandidate) == 0 {
		return &falseResult, nil
	}

	for _, project := range worklog.Projects {
		if userNameCandidate == project.Key {
			return &falseResult, nil
		}
	}

	return &trueResult, nil
}

func (s *ExcelService) finalize(f *excelize.File, sheet, dateFrom, dateTo string, context *models.ExcelContext) error {
	// dates
	startDate, err := time.Parse(constants.InputDateFormat, dateFrom)
	if err != nil {
		return err
	}

	endDate, err := time.Parse(constants.InputDateFormat, dateTo)
	if err != nil {
		return err
	}

	borders := []excelize.Border{
		{Type: "top", Color: "#000000", Style: 1},
		{Type: "left", Color: "#000000", Style: 1},
		{Type: "bottom", Color: "#000000", Style: 1},
		{Type: "right", Color: "#000000", Style: 1},
	}
	alignment := excelize.Alignment{Horizontal: "center", Vertical: "center"}
	font := excelize.Font{Size: 12, Color: "#000000", Bold: true}
	fill := excelize.Fill{Color: []string{"#FEC7CE"}, Type: "pattern", Pattern: 3}
	headerStyle, err := f.NewStyle(&excelize.Style{Border: borders, Alignment: &alignment, Font: &font, Fill: fill})
	if err != nil {
		return err
	}

	bodyStyle, err := f.NewStyle(&excelize.Style{Fill: fill})
	if err != nil {
		return err
	}

	i := 6
	for date := startDate; !date.After(endDate); date = date.AddDate(0, 0, 1) {
		i++
		weekday := date.Weekday()

		if weekday == 0 || weekday == 6 {
			headerCell, err := excelize.CoordinatesToCellName(i, 1)
			if err != nil {
				return err
			}

			err = f.SetCellStyle(sheet, headerCell, headerCell, headerStyle)
			if err != nil {
				return err
			}

			bodyTopCell, err := excelize.CoordinatesToCellName(i, 2)
			if err != nil {
				return err
			}

			bodyBottomCell, err := excelize.CoordinatesToCellName(i, context.LastRowIndex-1)
			if err != nil {
				return err
			}

			err = f.SetCellStyle(sheet, bodyTopCell, bodyBottomCell, bodyStyle)
			if err != nil {
				return err
			}
		}
	}

	return nil
}
