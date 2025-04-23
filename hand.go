package main

import (
	"github.com/tealeg/xlsx"
	"log"
	"os"
	"strconv"
	"strings"
)

type Employee struct {
	Apartment   string
	ID          string
	Name        string
	QuoterAward int
	Award       int
	Note        string
}

func main() {
	var isQuoter bool
	files, err := os.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}
	var allEmployees []Employee
	for _, file := range files {
		if !strings.HasSuffix(file.Name(), ".xlsx") && !strings.HasSuffix(file.Name(), ".xls") {
			continue
		}
		employees, quoter, err := readExcelFile(file.Name())
		isQuoter = quoter
		if err != nil {
			log.Printf("Error reading file %s: %v", file.Name(), err)
			continue // 或者根据需要返回错误或处理错误
		}

		allEmployees = append(allEmployees, employees...)
	}
	newEmployees := sumAwards(allEmployees, isQuoter)
	createNewFile(newEmployees, isQuoter)

}

func readExcelFile(filePath string) ([]Employee, bool, error) {
	var employees []Employee
	var isQuoter bool
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		return nil, false, err
	}

	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows[1:] {
			var employee Employee
			if len(row.Cells) < 5 { // 确保至少有5列数据
				continue
			} else if len(row.Cells) == 5 {
				isQuoter = false
				award, _ := row.Cells[3].Int()
				employee = Employee{
					Apartment: row.Cells[0].String(), //部门
					ID:        row.Cells[1].String(), // 员工编号
					Name:      row.Cells[2].String(), // 姓名
					Award:     award,                 // 奖金
					Note:      row.Cells[4].String(),
				}
			} else {
				isQuoter = true
				quoterAward, _ := row.Cells[3].Int()
				award, _ := row.Cells[4].Int()
				employee = Employee{
					Apartment:   row.Cells[0].String(), //部门
					ID:          row.Cells[1].String(), // 员工编号
					Name:        row.Cells[2].String(), // 姓名
					QuoterAward: quoterAward,           //季度奖金
					Award:       award,                 // 奖金
					Note:        row.Cells[5].String(),
				}
			}
			employees = append(employees, employee)
		}
	}
	return employees, isQuoter, nil
}

func sumAwards(allEmployees []Employee, isQuoter bool) []Employee {
	employeeMap := make(map[string]Employee)
	for _, employee := range allEmployees {
		if isQuoter {
			if existingEmployee, ok := employeeMap[employee.ID]; ok {
				existingEmployee.Award += employee.Award
				existingEmployee.QuoterAward += employee.QuoterAward
				employeeMap[employee.ID] = existingEmployee
			} else {
				employeeMap[employee.ID] = employee
			}
		} else {
			if existingEmployee, ok := employeeMap[employee.ID]; ok {
				existingEmployee.Award += employee.Award
				employeeMap[employee.ID] = existingEmployee
			} else {
				employeeMap[employee.ID] = employee
			}
		}
	}
	var result []Employee
	for _, employee := range employeeMap {
		result = append(result, employee)
	}
	return result
}

func createNewFile(newAllEmployees []Employee, isQuoter bool) {
	newXlFile := xlsx.NewFile()
	sheet, err := newXlFile.AddSheet("Sheet1")
	if err != nil {
		log.Fatalf("Error adding sheet: %v", err)
	}
	row := sheet.AddRow()
	if isQuoter {
		row.AddCell().Value = "部门"
		row.AddCell().Value = "员工编号"
		row.AddCell().Value = "姓名"
		row.AddCell().Value = "季度奖金"
		row.AddCell().Value = "奖金"
		row.AddCell().Value = "备注"
		for _, emp := range newAllEmployees {
			row = sheet.AddRow()
			row.AddCell().Value = emp.Apartment
			row.AddCell().Value = emp.ID
			row.AddCell().Value = emp.Name
			row.AddCell().Value = strconv.Itoa(emp.QuoterAward)
			row.AddCell().Value = strconv.Itoa(emp.Award)
			row.AddCell().Value = emp.Note
		}
	} else {
		cell0 := row.AddCell()
		cell0.Value = "部门"
		cell1 := row.AddCell()
		cell1.Value = "员工编号"
		cell2 := row.AddCell()
		cell2.Value = "姓名"
		cell3 := row.AddCell()
		cell3.Value = "月度奖金"
		cell4 := row.AddCell()
		cell4.Value = "备注"
		for _, emp := range newAllEmployees {
			row = sheet.AddRow()
			cell0 = row.AddCell()
			cell0.Value = emp.Apartment
			cell1 = row.AddCell()
			cell1.Value = emp.ID
			cell2 = row.AddCell()
			cell2.Value = emp.Name
			cell3 = row.AddCell()
			cell3.Value = strconv.Itoa(emp.Award)
			cell4 = row.AddCell()
			cell4.Value = emp.Note
		}
	}
	err = newXlFile.Save("综合.xlsx")
	if err != nil {
		log.Fatalf("Error saving file: %v", err)
	}
}
