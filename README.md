# Sheety

Easily load csv or xlsx formats into a slice of structs

## Installation

go get -u github.com/chriswalz/Sheety

## Usage
```go
    package main

    import (
    	"github.com/chriswalz/sheety"
    )

    type Student struct {
    	ID    float64
    	Grade float64
    }

    func main()  {
        s, err := sheety.OpenCSV("grades.csv")
    	if err != nil {
    		panic(err)
    	}
     
        // use pointer to struct
    	students := make([]*Student, 0, 3)
    
    	// map column index to struct field name
    	// First index starts at 1 NOT 0
    	err = s.ReadRows(1, &students, map[int]string{
    		1: "ID",
    		2: "Grade",
    	})
    	for _, v := range students {
    		// do work 
    	}
   }
```


See `sheety_test.go` for example usage 

