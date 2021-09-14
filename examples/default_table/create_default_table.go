package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"strconv"

	excel "github.com/I0HuKc/go-excel"
)

/* Structures for test data */

type Address struct {
	City    string `json:"city"`
	Street  string `json:"string"`
	Number  int    `json:"number"`
	Zipcode string `json:"zipcode"`
}

type Name struct {
	Firstname string `json:"firstname"`
	LastName  string `json:"lastname"`
}

type User struct {
	Id       int     `json:"id"`
	Address  Address `json:"address"`
	Email    string  `json:"email"`
	Username string  `json:"username"`
	Password string  `json:"password"`
	Name     Name    `json:"name"`
	Phone    string  `json:"phone"`
}

/* Table parameters */

// The line from which the table will begin
var tableStartFromLine = 1

// Column names
var headValues = []string{"ID", "Name", "Username", "Email", "Phone", "Password", "City", "Zipcode"}

// Header parameters
var tableHeader excel.Header = excel.Header{
	CellParams: excel.CreateHeaderCell(headValues, strconv.Itoa(tableStartFromLine)),
	ColParams: []excel.ColWidth{
		{
			StartCol: "A",
			EndCol:   "A",
			Width:    10,
		},
		{
			StartCol: "B",
			EndCol:   "F",
			Width:    30,
		},
		{
			StartCol: "G",
			EndCol:   "H",
			Width:    25,
		},
	},
	RowParams: []excel.RowHeight{
		{
			Row:    tableStartFromLine,
			Height: 25,
		},
	},
}

func main() {
	/* Getting test data */
	resp, err := http.Get("https://fakestoreapi.com/users")
	if err != nil {
		fmt.Println(err)
	}
	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		fmt.Println(err)
	}
	data := &[]User{}
	if err = json.Unmarshal(body, data); err != nil {
		fmt.Println(err)
	}

	/* Create data array */
	var tableValue [][]interface{}
	for _, v := range *data {
		tableValue = append(tableValue, []interface{}{
			v.Id,
			fmt.Sprintf("%s %s", v.Name.Firstname, v.Name.LastName),
			v.Username,
			v.Email,
			v.Phone,
			v.Password,
			v.Address.City,
			v.Address.Zipcode,
		})
	}

	// Create a new xlsx file
	if err = excel.NewFile("users.xlsx"); err != nil {
		fmt.Println(err)
	}

	// Create default table
	if err = excel.CreateDefaultTable(excel.DefaultTable{
		PathName:         "users.xlsx",
		TableHeader:      tableHeader,
		Data:             tableValue,
		Sheet:            "Users",
		ContentRowHeight: 18,
		ContentLineStart: tableStartFromLine + 1,
	}); err != nil {
		fmt.Println(err)
	}
}
