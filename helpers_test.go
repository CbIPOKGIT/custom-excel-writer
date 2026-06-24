package customexcelwriter

import (
	"strings"
	"testing"
)

func TestPrepareSheetName(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{
			name:     "Normal sheet name",
			input:    "Monthly Sales",
			expected: "Monthly Sales",
		},
		{
			name:     "Replace dots with underscores",
			input:    "reports.2026.05",
			expected: "reports_2026_05",
		},
		{
			name:     "Replace forbidden characters",
			input:    "sales:Q1/Q2\\summary?*[charts]",
			expected: "sales_Q1_Q2_summary___charts_",
		},
		{
			name:     "Trim spaces and single quotes at ends",
			input:    "  'My Sheet Name'  ",
			expected: "My Sheet Name",
		},
		{
			name:     "Empty input fallback",
			input:    "",
			expected: "Sheet",
		},
		{
			name:     "Only forbidden/trimmed characters fallback",
			input:    " :///\\*?[]' ' ",
			expected: "Sheet",
		},
		{
			name:     "Truncate simple ASCII to 31 chars",
			input:    "ThisIsAVeryLongSheetNameThatExceedsThirtyOneCharactersLimit",
			expected: "ThisIsAVeryLongSheetNameThatExc",
		},
		{
			name:     "Truncate Cyrillic UTF-8 name to 31 runes",
			input:    "ДужеДовгаНазваЛистаЯкаПеревищуєТридцятьОдинСимвол",
			expected: "ДужеДовгаНазваЛистаЯкаПеревищує",
		},
		{
			name:     "Truncate and trim trailing quotes/spaces",
			input:    "SheetNameEndingInQuotesAfterTruncation      '   ",
			expected: "SheetNameEndingInQuotesAfterTru",
		},
		{
			name:     "Mixed replacement, trimming, and truncation",
			input:    "  'Sales.Report.For.Region.North:East/West?*[]'  ",
			expected: "Sales_Report_For_Region_North_E",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			actual := PrepareSheetName(tt.input)
			if actual != tt.expected {
				t.Errorf("PrepareSheetName(%q) = %q; expected %q", tt.input, actual, tt.expected)
			}
			// Verify resulting Excel constraints
			if len([]rune(actual)) > 31 {
				t.Errorf("Resulting name %q exceeds 31 characters limit", actual)
			}
			if actual == "" {
				t.Errorf("Resulting name cannot be empty")
			}
			for _, forbidden := range []string{"\\", "/", "?", "*", "[", "]", ":"} {
				if strings.Contains(actual, forbidden) {
					t.Errorf("Resulting name %q contains forbidden character %q", actual, forbidden)
				}
			}
			if strings.HasPrefix(actual, "'") || strings.HasSuffix(actual, "'") {
				t.Errorf("Resulting name %q starts or ends with a single quote", actual)
			}
		})
	}
}
