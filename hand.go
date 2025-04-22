package main

import (
	"github.com/tealeg/xlsx"
	"log"
	"os"
	"strconv"
	"strings"
)

type Employee struct {
	ID    string
	Name  string
	Award int
}

func main() {
	files, err := os.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}
	var allEmployees []Employee
	for _, file := range files {
		if !strings.HasSuffix(file.Name(), ".xlsx") {
			continue
		}
		employees, err := readExcelFile(file.Name())
		if err != nil {
			log.Printf("Error reading file %s: %v", file.Name(), err)
			continue // 或者根据需要返回错误或处理错误
		}

		allEmployees = append(allEmployees, employees...)
	}
	newEmployees := sumAwards(allEmployees)
	createNewFile(newEmployees)

}

func readExcelFile(filePath string) ([]Employee, error) {
	var employees []Employee
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows[1:] {
			if len(row.Cells) < 3 { // 确保至少有3列数据
				continue
			}
			award, _ := row.Cells[2].Int()
			employee := Employee{
				ID:    row.Cells[0].String(), // 假设员工编号在第一列
				Name:  row.Cells[1].String(), // 假设姓名在第二列
				Award: award,                 // 假设奖在第三列
			}
			employees = append(employees, employee)
		}
	}
	return employees, nil
}

func sumAwards(allEmployees []Employee) []Employee {
	employeeMap := make(map[string]Employee)
	for _, employee := range allEmployees {
		if existingEmployee, ok := employeeMap[employee.ID]; ok {
			existingEmployee.Award += employee.Award
			employeeMap[employee.ID] = existingEmployee
		} else {
			employeeMap[employee.ID] = employee
		}
	}
	var result []Employee
	for _, employee := range employeeMap {
		result = append(result, employee)
	}
	return result
}

func createNewFile(newAllEmployees []Employee) {
	newXlFile := xlsx.NewFile()
	sheet, err := newXlFile.AddSheet("Sheet1")
	if err != nil {
		log.Fatalf("Error adding sheet: %v", err)
	}
	row := sheet.AddRow()
	cell0 := row.AddCell()
	cell0.Value = "员工编号"
	cell1 := row.AddCell()
	cell1.Value = "姓名"
	cell2 := row.AddCell()
	cell2.Value = "月度奖金"
	for _, emp := range newAllEmployees {
		row = sheet.AddRow()
		cell0 = row.AddCell()
		cell0.Value = emp.ID
		cell1 = row.AddCell()
		cell1.Value = emp.Name
		cell2 = row.AddCell()
		cell2.Value = strconv.Itoa(emp.Award)
	}
	err = newXlFile.Save("new_example.xlsx")
	if err != nil {
		log.Fatalf("Error saving file: %v", err)
	}
}

