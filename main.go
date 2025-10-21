package main

import (
	"flag"
	"fmt"
	"log"
	"math/rand"
	"sort"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

var rnd *rand.Rand
var rnd2 *rand.Rand

type Volunteer struct {
	Name      string
	Email     string
	Phone     string
	Teacher   string
	EventType string // "party" or "fieldtrip"
	EventName string // specific party or field trip name
	Alternate bool
}

type EventConfig struct {
	Name  string
	Count int // Same count for all teachers
}

type PartyConfig EventConfig

type FieldTripConfig struct {
	EventConfig
	Teachers []string
}

func main() {
	inputFile := flag.String("input", "input.xlsx", "Input XLSX file")
	outputFile := flag.String("output", "output.xlsx", "Output XLSX file")
	seed := flag.Int64("seed", 42, "Random seed")
	flag.Parse()

	rnd = rand.New(rand.NewSource(*seed))
	rnd2 = rand.New(rand.NewSource(*seed * 2))

	volunteers, parties, fieldTrips, allTeachers, err := readInput(*inputFile)
	if err != nil {
		log.Fatalf("Error reading input: %v", err)
	}

	assignments := assignVolunteers(volunteers, parties, fieldTrips, allTeachers)

	if err := writeOutput(*outputFile, assignments, parties, fieldTrips); err != nil {
		log.Fatalf("Error writing output: %v", err)
	}

	fmt.Printf("Successfully created %s\n", *outputFile)
}

func readInput(filename string) ([]Volunteer, []PartyConfig, []FieldTripConfig, []string, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, nil, nil, nil, err
	}
	defer f.Close()

	// Read variables first to get party and field trip names
	parties, fieldTrips, err := readVariables(f)
	if err != nil {
		return nil, nil, nil, nil, err
	}

	// Extract party and field trip names for parsing
	partyNames := make([]string, len(parties))
	for i, p := range parties {
		partyNames[i] = p.Name
	}

	tripNames := make([]string, len(fieldTrips))
	for i, t := range fieldTrips {
		tripNames[i] = t.Name
	}

	volunteers, allTeachers, err := readVolunteersWithNames(f, partyNames, tripNames)
	if err != nil {
		return nil, nil, nil, nil, err
	}

	return volunteers, parties, fieldTrips, allTeachers, nil
}

func readVolunteersWithNames(f *excelize.File, partyNames, tripNames []string) ([]Volunteer, []string, error) {
	rows, err := f.GetRows("Form Responses 1")
	if err != nil {
		return nil, nil, err
	}

	var volunteers []Volunteer
	teacherSet := make(map[string]bool)
	indexes := map[string]int{}
	for i, row := range rows {
		if i == 0 {
			for id, _str := range row {
				str := strings.ToLower(_str)
				if strings.Contains(str, "teacher") {
					indexes["teacher"] = id
				} else if strings.Contains(str, "first and last name") {
					indexes["name"] = id
				} else if strings.Contains(str, "first name") {
					indexes["firstName"] = id
				} else if strings.Contains(str, "last name") {
					indexes["lastName"] = id
				} else if strings.Contains(str, "phone") {
					indexes["phone"] = id
				} else if strings.Contains(str, "email") {
					indexes["email"] = id
				} else if strings.Contains(str, "party or parties") {
					indexes["parties"] = id
				} else if strings.Contains(str, "field trip(s)") {
					indexes["fieldtrips"] = id
				}
			}
		}
		if i == 0 || len(row) < 6 {
			continue // Skip header
		}

		teacher := strings.TrimSpace(row[indexes["teacher"]])
		var name string
		if nid, ok := indexes["name"]; ok {
			name = strings.TrimSpace(row[nid])
		} else {
			name = strings.TrimSpace(row[indexes["firstName"]]) + " " + strings.TrimSpace(row[indexes["lastName"]])
		}

		phone := strings.TrimSpace(row[indexes["phone"]])
		email := strings.TrimSpace(row[indexes["email"]])

		teacherSet[teacher] = true

		partiesStr := ""
		fieldTripsStr := ""

		if len(row) > indexes["parties"] {
			partiesStr = row[indexes["parties"]]
		}
		if len(row) > indexes["fieldtrips"] {
			fieldTripsStr = row[indexes["fieldtrips"]]
		}

		// Parse parties using known party names
		if partiesStr != "" && !strings.Contains(partiesStr, "N/A") {
			for _, partyName := range partyNames {
				if strings.Contains(partiesStr, partyName) {
					volunteers = append(volunteers, Volunteer{
						Name:      name,
						Phone:     phone,
						Email:     email,
						Teacher:   teacher,
						EventType: "party",
						EventName: partyName,
					})
				}
			}
		}

		// Parse field trips using known trip names
		if fieldTripsStr != "" && !strings.Contains(fieldTripsStr, "N/A") {
			for _, tripName := range tripNames {
				if strings.Contains(fieldTripsStr, tripName) {
					volunteers = append(volunteers, Volunteer{
						Name:      name,
						Phone:     phone,
						Email:     email,
						Teacher:   teacher,
						EventType: "fieldtrip",
						EventName: tripName,
					})
				}
			}
		}
	}

	var allTeachers []string
	for teacher := range teacherSet {
		allTeachers = append(allTeachers, teacher)
	}

	return volunteers, allTeachers, nil
}

func readVariables(f *excelize.File) ([]PartyConfig, []FieldTripConfig, error) {
	rows, err := f.GetRows("Variables")
	if err != nil {
		return nil, nil, err
	}

	var parties []PartyConfig
	var fieldTrips []FieldTripConfig

	for i, row := range rows {
		if i == 0 || len(row) < 5 {
			continue // Skip header
		}

		// Parties: Col A (Party Name), Col B (Count Per Teacher)
		if row[0] != "" && row[1] != "" {
			parties = append(parties, PartyConfig{
				Name:  strings.TrimSpace(row[0]),
				Count: parseInt(row[1]),
			})
		}

		// Field Trips: Col C (Trip Name), Col D (Teachers pipe-separated), Col E (Count)
		if row[2] != "" && row[3] != "" && row[4] != "" {
			teachersStr := strings.TrimSpace(row[3])
			var teachers []string

			if teachersStr == "ALL" {
				teachers = []string{"ALL"} // Special marker for all teachers
			} else {
				for _, t := range strings.Split(teachersStr, "|") {
					t = strings.TrimSpace(t)
					if t != "" {
						teachers = append(teachers, t)
					}
				}
			}

			fieldTrips = append(fieldTrips, FieldTripConfig{
				EventConfig: EventConfig{
					Name:  strings.TrimSpace(row[2]),
					Count: parseInt(row[4]),
				},
				Teachers: teachers,
			})
		}
	}

	return parties, fieldTrips, nil
}

func assignVolunteers(volunteers []Volunteer, parties []PartyConfig, fieldTrips []FieldTripConfig, allTeachers []string) map[string]map[string][]Volunteer {
	assignments := make(map[string]map[string][]Volunteer)
	usedForParty := make(map[string]bool)
	usedForFieldTrip := make(map[string]bool)
	usedForThisEvent := make(map[string]map[string]bool)

	// Assign parties (all teachers get same count)
	for _, party := range parties {
		assignments[party.Name] = make(map[string][]Volunteer)
		usedForThisEvent[party.Name] = make(map[string]bool)

		// Get all volunteers who signed up for this party
		candidates := filterVolunteersByEvent(volunteers, "party", party.Name)
		rnd.Shuffle(len(candidates), func(i, j int) {
			candidates[i], candidates[j] = candidates[j], candidates[i]
		})

		// For each teacher, assign volunteers
		for _, teacher := range allTeachers {
			teacherCandidates := filterByTeacher(candidates, teacher)
			assigned := 0

			for _, v := range teacherCandidates {
				if assigned >= party.Count {
					break
				}
				name := nameKey(v.Name)
				if !usedForParty[name] {
					assignments[party.Name][teacher] = append(assignments[party.Name][teacher], v)
					usedForParty[name] = true
					usedForThisEvent[party.Name][name] = true
					assigned++
				}
			}
		}
	}

	// Assign field trips
	for _, trip := range fieldTrips {
		assignments[trip.Name] = make(map[string][]Volunteer)
		usedForThisEvent[trip.Name] = make(map[string]bool)

		candidates := filterVolunteersByEvent(volunteers, "fieldtrip", trip.Name)

		rnd.Shuffle(len(candidates), func(i, j int) {
			candidates[i], candidates[j] = candidates[j], candidates[i]
		})

		teachers := trip.Teachers
		if len(teachers) == 1 && teachers[0] == "ALL" {
			teachers = allTeachers
		}

		for _, teacher := range teachers {
			teacherCandidates := filterByTeacher(candidates, teacher)
			assigned := 0

			for _, v := range teacherCandidates {
				if assigned >= trip.Count {
					break
				}
				name := nameKey(v.Name)
				if !usedForFieldTrip[name] {
					assignments[trip.Name][teacher] = append(assignments[trip.Name][teacher], v)
					usedForFieldTrip[name] = true
					usedForThisEvent[trip.Name][name] = true

					assigned++
				}
			}
		}
	}

	// Add 2 alternates for each party/teacher
	usedAsAlternate := make(map[string]bool)
	for _, party := range parties {
		candidates := filterVolunteersByEvent(volunteers, "party", party.Name)
		rnd2.Shuffle(len(candidates), func(i, j int) {
			candidates[i], candidates[j] = candidates[j], candidates[i]
		})

		for _, teacher := range allTeachers {
			teacherCandidates := filterByTeacher(candidates, teacher)
			assigned := 0
			unique := true

		alternateParty:
			for _, v := range teacherCandidates {
				if assigned >= 2 {
					break
				}
				name := nameKey(v.Name)
				usedForThisEvent := usedForThisEvent[party.Name][name]
				// Skip if already assigned as primary or alternate
				if (!unique || !usedForParty[name]) && !usedForThisEvent && !usedAsAlternate[name] {
					v.Alternate = true
					assignments[party.Name][teacher] = append(assignments[party.Name][teacher], v)
					usedAsAlternate[name] = true
					assigned++
				}
			}
			if assigned < 2 && unique {
				unique = false
				goto alternateParty
			}
		}
	}

	// Add 2 alternates for each field trip/teacher
	usedAsAlternateTrip := make(map[string]bool)
	for _, trip := range fieldTrips {
		candidates := filterVolunteersByEvent(volunteers, "fieldtrip", trip.Name)
		rnd2.Shuffle(len(candidates), func(i, j int) {
			candidates[i], candidates[j] = candidates[j], candidates[i]
		})

		teachers := trip.Teachers
		if len(teachers) == 1 && teachers[0] == "ALL" {
			teachers = allTeachers
		}

		for _, teacher := range teachers {
			teacherCandidates := filterByTeacher(candidates, teacher)
			assigned := 0
			unique := true

		alternateTrip:
			for _, v := range teacherCandidates {
				if assigned >= 2 {
					break
				}
				name := nameKey(v.Name)
				usedForThisEvent := usedForThisEvent[trip.Name][name]
				// Skip if already assigned as primary or alternate
				if (!unique || !usedForFieldTrip[name]) && !usedForThisEvent && !usedAsAlternateTrip[name] {
					v.Alternate = true
					assignments[trip.Name][teacher] = append(assignments[trip.Name][teacher], v)
					usedAsAlternateTrip[name] = true
					assigned++
				}
			}
			if assigned < 2 && unique {
				unique = false
				goto alternateTrip
			}
		}
	}

	return assignments
}

func writeOutput(filename string, assignments map[string]map[string][]Volunteer, parties []PartyConfig, fieldTrips []FieldTripConfig) error {
	f := excelize.NewFile()

	index := 0

	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 22, Color: "2C3E50"},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"E8F4F8"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "bottom", Color: "B4C7E7", Style: 1},
		},
	})

	// Define styles with better padding and subtle colors
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 14, Color: "2C3E50"},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"E8F4F8"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "bottom", Color: "B4C7E7", Style: 1},
		},
	})

	teacherHeaderStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 12, Color: "1F4E78"},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"D9E2F3"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "bottom", Color: "8EA9DB", Style: 1},
			{Type: "top", Color: "8EA9DB", Style: 1},
			{Type: "left", Color: "8EA9DB", Style: 1},
			{Type: "right", Color: "8EA9DB", Style: 1},
		},
	})

	volunteerStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
	})

	alternateDividerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 12, Color: "7F7F7F"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"F2F2F2"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "BFBFBF", Style: 1},
			{Type: "bottom", Color: "BFBFBF", Style: 1},
		},
	})

	alternateStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Italic: true, Color: "595959"},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"F9F9F9"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "bottom", Color: "BFBFBF", Style: 1},
		},
	})

	// Write parties first, in order from Variables sheet
	for _, party := range parties {
		teacherAssignments := assignments[party.Name]
		sheetName := sanitizeSheetName(party.Name)

		if idx, _ := f.GetSheetIndex(sheetName); idx >= 0 {
			sheetName = sheetName[:len(sheetName)-2] + " 2"
		}

		if index == 0 {
			f.SetSheetName("Sheet1", sheetName)
		} else {
			f.NewSheet(sheetName)
		}

		writeEventSheet(f, sheetName, EventConfig(party), teacherAssignments, party.Count, titleStyle, headerStyle, teacherHeaderStyle, volunteerStyle, alternateDividerStyle, alternateStyle)

		index++
	}

	// Write field trips second, in order from Variables sheet
	for _, trip := range fieldTrips {
		teacherAssignments := assignments[trip.Name]
		sheetName := sanitizeSheetName(trip.Name)

		if idx, _ := f.GetSheetIndex(sheetName); idx >= 0 {
			sheetName = sheetName[:len(sheetName)-2] + " 2"
		}

		f.NewSheet(sheetName)

		writeEventSheet(f, sheetName, trip.EventConfig, teacherAssignments, trip.Count, titleStyle, headerStyle, teacherHeaderStyle, volunteerStyle, alternateDividerStyle, alternateStyle)

		index++
	}

	return f.SaveAs(filename)
}

func writeEventSheet(f *excelize.File, sheetName string, event EventConfig, teacherAssignments map[string][]Volunteer, primaryCount int, titleStyle, headerStyle, teacherHeaderStyle, volunteerStyle, alternateDividerStyle, alternateStyle int) {

	f.SetRowHeight(sheetName, 1, 36)
	f.SetCellValue(sheetName, "A1", event.Name)
	f.MergeCell(sheetName, "A1", "D1")
	f.SetCellStyle(sheetName, "A1", "D1", titleStyle)

	// Set column widths with more generous spacing
	f.SetColWidth(sheetName, "A", "A", 28)
	f.SetColWidth(sheetName, "B", "B", 32)
	f.SetColWidth(sheetName, "C", "C", 18)
	f.SetColWidth(sheetName, "D", "D", 35)

	// Headers
	f.SetCellValue(sheetName, "A2", "Teacher")
	f.SetCellValue(sheetName, "B2", "Name")
	f.SetCellValue(sheetName, "C2", "Phone")
	f.SetCellValue(sheetName, "D2", "Email")
	f.SetCellStyle(sheetName, "A2", "D2", headerStyle)
	f.SetRowHeight(sheetName, 2, 25)

	row := 3

	// Sort teachers for consistent output
	var sortedTeachers []string
	for teacher := range teacherAssignments {
		sortedTeachers = append(sortedTeachers, teacher)
	}

	sort.Strings(sortedTeachers)

	for _, teacher := range sortedTeachers {
		volunteers := teacherAssignments[teacher]

		// Teacher section header
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), teacher)
		f.MergeCell(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row))
		f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row), teacherHeaderStyle)
		f.SetRowHeight(sheetName, row, 22)
		row++

		// Primary volunteers
		prevAlternate := false
		for i := 0; i < len(volunteers); i++ {
			v := volunteers[i]

			if v.Alternate && !prevAlternate {
				// Divider row for alternates
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), "ALTERNATES")
				f.MergeCell(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row))
				f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row), alternateDividerStyle)
				f.SetRowHeight(sheetName, row, 20)
				row++
				prevAlternate = true
			}

			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), teacher)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), v.Name)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), FormatPhoneNumber(v.Phone))
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), v.Email)
			if !v.Alternate {
				f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row), volunteerStyle)
			} else {
				f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row), alternateStyle)

			}
			f.SetRowHeight(sheetName, row, 20)
			row++

			// Check to see if we are at the end and need more rows
			// or if we're moving on to the alternates
			if !prevAlternate && (len(volunteers) <= i+1 || volunteers[i+1].Alternate) {
				if i < event.Count-1 {
					addCnt := event.Count - i - 1
					f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row+addCnt), volunteerStyle)
					row += addCnt
				}
			}
		}

		// Spacing between teachers
		row++
	}
}

// Helper functions

func FormatPhoneNumber(input string) string {
	// Strip all non-digit characters
	var digits strings.Builder
	for _, char := range input {
		if unicode.IsDigit(char) {
			digits.WriteRune(char)
		}
	}

	phoneDigits := digits.String()

	// Check for exactly 10 digits
	if len(phoneDigits) != 10 {
		return input
	}

	// Format as (xxx) xxx-xxxx
	formatted := fmt.Sprintf("(%s) %s-%s",
		phoneDigits[0:3],
		phoneDigits[3:6],
		phoneDigits[6:10])

	return formatted
}

func filterVolunteersByEvent(volunteers []Volunteer, eventType, eventName string) []Volunteer {
	var filtered []Volunteer
	for _, v := range volunteers {
		if v.EventType == eventType && v.EventName == eventName {
			filtered = append(filtered, v)
		}
	}
	return filtered
}

func filterByTeacher(volunteers []Volunteer, teacher string) []Volunteer {
	var filtered []Volunteer
	for _, v := range volunteers {
		if v.Teacher == teacher {
			filtered = append(filtered, v)
		}
	}
	return filtered
}

func parseInt(s string) int {
	var i int
	fmt.Sscanf(s, "%d", &i)
	return i
}

func nameKey(name string) string {
	return strings.ToLower(name)
}
func sanitizeSheetName(name string) string {
	name = strings.ReplaceAll(name, "/", "-")
	name = strings.ReplaceAll(name, "\\", "-")
	name = strings.ReplaceAll(name, "?", "")
	name = strings.ReplaceAll(name, "*", "")
	name = strings.ReplaceAll(name, "[", "")
	name = strings.ReplaceAll(name, "]", "")
	name = strings.ReplaceAll(name, ":", "")
	if len(name) > 31 {
		name = name[:31]
	}
	return name
}
