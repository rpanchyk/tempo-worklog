# Functions
define make_build
    GOOS=$(1) GOARCH=$(2) go build -o builds/$(1)/$(2)/$(3)
	cp -f TemplateAppConfig.yaml builds/$(1)/$(2)/
	cd builds/$(1)/$(2) && rm -f tempo-worklog.zip && zip --recurse-paths --move tempo-worklog.zip . && cd -
endef

# Batch build
build: deps build-linux build-macosx build-windows

# Update dependecies
deps:
	go mod tidy && go mod vendor

# Linux
build-linux:
	$(call make_build,linux,amd64,tempo-worklog)

# MacOSX
build-macosx:
	$(call make_build,darwin,amd64,tempo-worklog)
	$(call make_build,darwin,arm64,tempo-worklog)

# Windows
build-windows:
	$(call make_build,windows,amd64,tempo-worklog.exe)
