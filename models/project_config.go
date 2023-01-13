package models

type ProjectConfigWrapper struct {
	ProjectKeyToConfig map[string]ProjectConfig
}

type ProjectConfig struct {
	UserNameToConfig map[string]UserConfig
}

type UserConfig struct {
	Position string
	Rate     int
}
