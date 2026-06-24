package customexcelwriter

import (
	"strings"
)

// PrepareSheetName sanitizes and prepares a sheet name according to Excel limits and requirements.
// Excel limits sheet names to 31 characters, and they cannot be blank or contain: \ / ? * [ ] :
// This function replaces those forbidden characters and '.' with '_', trims spaces and single quotes,
// and ensures the sheet name remains valid.
func PrepareSheetName(name string) string {
	// First, replace dots with underscores per user request
	name = strings.ReplaceAll(name, ".", "_")

	// Replace forbidden characters: \ / ? * [ ] : with underscores
	forbidden := []string{"\\", "/", "?", "*", "[", "]", ":"}
	for _, f := range forbidden {
		name = strings.ReplaceAll(name, f, "_")
	}

	// Trim leading and trailing spaces and single quotes
	name = strings.Trim(name, " '")

	// Limit length to 31 runes (unicode characters) to handle multibyte characters correctly
	runes := []rune(name)
	if len(runes) > 31 {
		name = string(runes[:31])
		// After truncation, trim any resulting leading/trailing spaces or single quotes
		name = strings.Trim(name, " '")
	}

	// If the name is empty or consists entirely of underscores after processing, use a fallback default name
	if strings.Trim(name, "_") == "" {
		return "Sheet"
	}

	return name
}