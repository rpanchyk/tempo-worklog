package services

import (
	"github.com/spf13/viper"
	"log"
	"tempo-worklog/models"
	"tempo-worklog/utils"
)

type AppConfigService struct {
	filePath string
}

func NewAppConfigService(filePath string) *AppConfigService {
	return &AppConfigService{filePath: filePath}
}

func (s *AppConfigService) Get() (*models.AppConfig, error) {
	viper.SetConfigFile(s.filePath)
	viper.SetConfigType("yaml")

	err := viper.ReadInConfig()
	if err != nil {
		return nil, err
	}

	var appConfig models.AppConfig

	err = viper.Unmarshal(&appConfig)
	if err != nil {
		return nil, err
	}

	log.Println("Parsed", s.filePath, utils.ToPrettyString("config", appConfig))

	return &appConfig, nil
}
