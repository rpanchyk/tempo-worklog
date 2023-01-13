package services

import (
	"errors"
	"log"
	"os"
	"strings"
	"tempo-worklog/constants"
	"tempo-worklog/models"
	"time"
)

type InputArgsService struct {
}

func NewInputArgsService() *InputArgsService {
	return &InputArgsService{}
}

func (s *InputArgsService) Parse(args []string) (*models.InputArgs, error) {
	if len(args) < 4 {
		return nil, errors.New("not enough input arguments")
	}

	// 1st
	configFile := args[0]
	_, err := os.Stat(configFile)
	if err != nil {
		return nil, err
	}
	log.Println("Validated config-file:", configFile)

	// 2nd
	rawProjects := args[1]
	var projects []string
	for _, project := range strings.Split(rawProjects, ",") {
		if len(project) > 0 {
			projects = append(projects, project)
		}
	}
	if len(projects) == 0 {
		return nil, errors.New("projects are not set")
	}
	log.Println("Validated projects:", strings.Join(projects[:], ", "))

	// 3rd
	dateFrom := args[2]
	_, err = time.Parse(constants.InputDateFormat, dateFrom)
	if err != nil {
		return nil, err
	}
	log.Println("Validated date-from:", dateFrom)

	// 4th
	dateTo := args[3]
	_, err = time.Parse(constants.InputDateFormat, dateTo)
	if err != nil {
		return nil, err
	}
	log.Println("Validated date-to:", dateTo)

	result := &models.InputArgs{
		ConfigFile: configFile,
		Projects:   projects,
		DateFrom:   dateFrom,
		DateTo:     dateTo,
	}

	return result, nil
}
