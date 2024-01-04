package exceltemplate

import (
	"testing"
	"fmt"
)

func TestReplaceString(t *testing.T) {
	orgStr := "%{greet} %{person.id}, my name is %{person.name}"
	data := map[string]interface{}{
		"greet":  "hello",
		"person": map[string]interface{}{
			"list":[]interface{}{
				map[string]interface{}{
					"id": 1, 
					"name": "zhangsan",
				},
			},
		},
	}
	repalcedStr,replaced := ReplaceString(orgStr, data)
	fmt.Println(replaced)
	fmt.Println(orgStr)
	fmt.Println(repalcedStr)
	if !replaced {
		t.Errorf("ReplaceString failed")
	}

	if repalcedStr != "hello 1, my name is zhangsan" {
		t.Errorf("ReplaceString failed")
	}
}

func TestGetExcelFromTemplate(t *testing.T){
	filename:=`./fortest.xlsx`
	data := map[string]interface{}{
		"greet":  "hello",
		"person": map[string]interface{}{
			"list":[]interface{}{
				map[string]interface{}{
					"id": 1, 
					"name": "zhangsan",
				},
			},
		},
	}
	f,err:=GetExcelFromTemplate(filename,data)
	if err != nil {
		fmt.Println(err)
		t.Errorf("GetExcelFromTemplate failed")
		return
	}
	defer f.Close()

	sheetName := "sheet1"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		t.Errorf("GetExcelFromTemplate failed")
		return
	}

	//遍历所有的行
	for _, row := range rows {
		//遍历行中的所有单元格
		for _, cellValue := range row {
			fmt.Println(cellValue)
			if cellValue != "hello 1, my name is zhangsan" {
				t.Errorf("GetExcelFromTemplate failed")
				return
			}
		}
	}
}