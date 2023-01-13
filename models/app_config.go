package models

type AppConfig struct {
	Jira  JiraAppConfig  `mapstructure:"jira"`
	Files FilesAppConfig `mapstructure:"files"`
}

type JiraAppConfig struct {
	Url        string `mapstructure:"url"`
	UserEmail  string `mapstructure:"user_email"`
	UserToken  string `mapstructure:"user_token"`
	TempoToken string `mapstructure:"tempo_token"`
}

type FilesAppConfig struct {
	ProjectConfigFile string `mapstructure:"project_config"`
	ReportFile        string `mapstructure:"report"`
}
