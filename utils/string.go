package utils

import (
	"encoding/json"
	"fmt"
	"log"
)

func ToPrettyString(prefix string, obj interface{}) string {
	pretty, err := json.MarshalIndent(obj, "", "  ")
	if err != nil {
		log.Fatal(err)
		return "error while prettifying"
	}
	return fmt.Sprintf("%s: \r\n%s", prefix, string(pretty))
}
