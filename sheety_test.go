package sheety

import (
	"testing"
	"fmt"
)

type Student struct {
	ID    float64
	Grade float64
}

var studentsExpected = []Student{
	{
		ID:    43875,
		Grade: 84,
	},
	{
		ID:    20347,
		Grade: 72,
	},
	{
		ID:    83274,
		Grade: 83,
	},
	{
		ID:    72345,
		Grade: 99,
	},
}

func TestXLSX(t *testing.T) {
	s, err := OpenSpreadsheet("grades.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	students := make([]*Student, 0, 3)
	s.ReadRows(3, &students, map[int]string{
		1: "ID",
		2: "Grade",
	})
	for i, expected := range studentsExpected {
		got := students[i]
		if expected.ID != got.ID {
			t.Fatalf("length: exp=%v, got=%v", expected.ID, got.ID)
		}
		if expected.Grade != got.Grade {
			t.Fatal(fmt.Sprintf("exp=%v, got=%v", expected.Grade, got.Grade))
		}
	}
}
func TestCSV(t *testing.T) {
	s, err := OpenCSV("grades.csv")
	if err != nil {
		t.Fatal(err)
	}
	students := make([]*Student, 0, 3)

	err = s.ReadRows(1, &students, map[int]string{
		1: "ID",
		2: "Grade",
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(students) > 4 {
		t.Fatalf("exp=%v got=%v", 4, len(students))
	}
	for i, expected := range studentsExpected {
		got := students[i]
		if expected.ID != got.ID {
			t.Fatalf("length: exp=%v, got=%v", expected.ID, got.ID)
		}
		if expected.Grade != got.Grade {
			t.Fatal(fmt.Sprintf("exp=%v, got=%v", expected.Grade, got.Grade))
		}
	}
}
