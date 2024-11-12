// ----------------------------------------------------------------------------
//
// Spreadsheet support functions
//
// Author: William Shaffer
// Version: 24-Sep-2024
//
// Copyright (c) William Shaffer
//
// ----------------------------------------------------------------------------

package spreadsheet

// ----------------------------------------------------------------------------
// Imports
// ----------------------------------------------------------------------------

import (
	"strconv"

	dec "github.com/shopspring/decimal"
	d "github.com/waysys/waydate/pkg/date"
)

// ----------------------------------------------------------------------------
// Types
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
// Functions
// ----------------------------------------------------------------------------

// cellName generates a string representing a cell in the spreadsheet.
func CellName(column string, row int) string {
	var cellName = column + strconv.Itoa(row)
	return cellName
}

// writeCell outputs a string value to the specified cell
func WriteCell(
	outputPtr *SpreadsheetFile,
	column string,
	row int,
	value string) {

	var cell = CellName(column, row)
	var err = outputPtr.SetCell(cell, value)
	Check(err, "Error writing cell "+cell+": ")
}

// writeCellInt outputs an integer value to the specified cell
func WriteCellInt(
	outputPtr *SpreadsheetFile,
	column string,
	row int,
	value int) {

	var cell = CellName(column, row)
	var err = outputPtr.SetCellInt(cell, value)
	Check(err, "Error writing cell "+cell+": ")
}

// writeCellFloat outputs a float64 value to the specified cell
func WriteCellFloat(
	outputPtr *SpreadsheetFile,
	column string,
	row int,
	value float64) {

	var cell = CellName(column, row)
	var err = outputPtr.SetCellFloat(cell, value)
	Check(err, "Error writing cell "+cell+": ")
}

// writeDecimal outputs a decimal value to the specified cell
func WriteCellDecimal(
	outputPtr *SpreadsheetFile,
	column string,
	row int,
	value dec.Decimal) {

	var cell = CellName(column, row)
	var err = outputPtr.SetCellDecimal(cell, value, FormatMoney)
	Check(err, "Error writing cell "+cell+": ")
}

// WriteCellDate outputs a WayDate to the specified cell.
func WriteCellDate(
	outputPtr *SpreadsheetFile,
	column string,
	row int,
	date d.Date) {
	var cell = CellName(column, row)
	var err = outputPtr.SetCellDate(cell, date)
	Check(err, "Error writing cell "+cell+": ")
}
