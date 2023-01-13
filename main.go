package main

import (
	"log"
	"os"
	"tempo-worklog/services"
)

func main() {
	log.Println("Report creating started")

	// args
	inputArgsService := services.NewInputArgsService()

	inputArgs, err := inputArgsService.Parse(os.Args[1:])
	if err != nil {
		log.Fatal(err)
		return
	}

	// config
	appConfigService := services.NewAppConfigService(inputArgs.ConfigFile)

	appConfig, err := appConfigService.Get()
	if err != nil {
		log.Fatal(err)
		return
	}

	// get data
	worklogService := services.NewWorklogService(
		appConfig.Jira.Url,
		appConfig.Jira.UserEmail,
		appConfig.Jira.UserToken,
		appConfig.Jira.TempoToken,
		services.NewProjectConfigService(appConfig.Files.ProjectConfigFile))

	worklog, err := worklogService.GetWorklog(inputArgs.Projects, inputArgs.DateFrom, inputArgs.DateTo)
	if err != nil {
		log.Fatal(err)
		return
	}

	// save data
	excelService := services.NewExcelService(appConfig.Files.ReportFile)

	err = excelService.Save(worklog, inputArgs.DateFrom, inputArgs.DateTo)
	if err != nil {
		log.Fatal(err)
		return
	}

	log.Println("Report creating finished successfully")
	log.Println("See", appConfig.Files.ReportFile)
}
