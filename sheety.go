package sheety

import (
	"errors"
	"github.com/tealeg/xlsx"
	"reflect"
	"log"
	"strconv"
	"fmt"
	"encoding/csv"
	"os"
	"bufio"
	"io"
)

type Spreadsheet struct {
	file *xlsx.File

}

type SpreadCSV struct {
	file *csv.Reader

}

func OpenCSV(excelFileName string) (*SpreadCSV, error) {
	f, err := os.Open(excelFileName)
	if err != nil {
		return nil, err
	}
	r := csv.NewReader(bufio.NewReader(f))
	r.FieldsPerRecord = -1

	return &SpreadCSV{
		file: r,
	}, nil
}

func OpenSpreadsheet(excelFileName string) (*Spreadsheet, error) {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return nil, err
	}

	return &Spreadsheet{
		file: xlFile,
	}, nil
}
// ReadRows takes a starting row index, a pointer to slice, and a mapping from spreadsheet to struct
func (s *Spreadsheet) ReadRows(start int, bb interface{}, m map[int]string) {
	if reflect.TypeOf(bb).Elem().Kind() != reflect.Slice {
		log.Println("nil szn")
		return
	}

	tElm := reflect.TypeOf(bb).Elem().Elem().Elem()
	//log.Println("kind: ", tElm.Kind())

	vv := reflect.ValueOf(bb).Elem()

	//reflect.Indirect()

	for _, sheet := range s.file.Sheets {
		for i, row := range sheet.Rows {
			if i < start {
				continue
			}
			vptr := reflect.New(tElm)
			for j, cell := range row.Cells {
				text := cell.String()
				name := m[j+1]
				if name == "" || text == ""{
					continue
				}
				//log.Println("Name:", name, "i: ", i)
				val := vptr.Elem().FieldByName(name)
				if val.Kind() == reflect.String {
					val.SetString(text)
				} else if val.Kind() == reflect.Float64 {
					num, err := strconv.ParseFloat(text, 64)
					if err != nil {
						panic(fmt.Sprintf("Text: %v, i: %v, j: %v", text, i, j))
					}
					val.SetFloat(num)
					//cell.Value = strconv.FormatFloat(val.Float(), 'f', 6, 64)
				} else {
					panic("type not implemented for parsing yet")
				}
			}
			vv.Set(reflect.Append(vv, vptr))
		}
	}
}
// ReadRows takes a starting row index, a pointer to slice, and a mapping from spreadsheet to struct
func (s *Spreadsheet) SaveRows(start int, bb interface{}, m map[int]string, path string) {
	if reflect.TypeOf(bb).Elem().Kind() != reflect.Slice {
		log.Println("nil szn")
		return
	}

	vv := reflect.ValueOf(bb).Elem()

	for _, sheet := range s.file.Sheets {
		for i, row := range sheet.Rows {
			if i < start {
				continue
			}
			if i >= vv.Len() {
				break
			}
			for col, name := range m {
				if col > len(row.Cells) {
					for x := 0; x <= col + 10; x++ {
						row.AddCell()
					}
				}
				//log.Println(len(row.Cells))
				//log.Println(col)
				cell := row.Cells[col]

				val := vv.Index(i).Elem().FieldByName(name)
				if val.Kind() == reflect.String {
					cell.Value = val.String()
				} else if val.Kind() == reflect.Float64 {
					cell.Value = strconv.FormatFloat(val.Float(), 'f', 6, 64)
				}
			}



		}
	}
	s.file.Save(path)
}

// ReadRows takes a starting row index, a pointer to slice, and a mapping from spreadsheet to struct
func (s *SpreadCSV) ReadRows(startIndex int, bb interface{}, m map[int]string) error {
	if reflect.TypeOf(bb).Elem().Kind() != reflect.Slice {
		log.Println("nil")
		return nil
	}

	tElm := reflect.TypeOf(bb).Elem().Elem().Elem()

	vv := reflect.ValueOf(bb).Elem()

	i := -1
	for {
		i++
		record, err := s.file.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if i < startIndex {
			continue
		}
		vptr := reflect.New(tElm)
		for j, text := range record {
			name := m[j+1]
			if name == "" || text == ""{
				continue
			}
			val := vptr.Elem().FieldByName(name)
			if val.Kind() == reflect.String {
				val.SetString(text)
			} else if val.Kind() == reflect.Float64 {
				num, err := strconv.ParseFloat(text, 64)
				if err != nil {
					return errors.New(fmt.Sprintf("Text: %v, i: %v, j: %v", text, i, j))
				}
				val.SetFloat(num)
			} else {
				return errors.New("type not implemented for parsing yet - contact maintainer or submit PR")
			}
		}
		vv.Set(reflect.Append(vv, vptr))
	}
	return nil
}
