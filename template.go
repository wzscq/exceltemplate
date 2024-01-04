package exceltemplate

import (
	"github.com/xuri/excelize/v2"
	"regexp"
	"fmt"
	"strings"
	"reflect"
)

func GetExcelFromTemplate(templateFile string, data map[string]interface{}) (*excelize.File, error) {
	//读取tempalteFile
	f, err := excelize.OpenFile(templateFile)
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	//遍历所有的sheet
	for _, sheetName := range f.GetSheetList() {
		//处理单个sheet
		err = processSheet(f, sheetName, data)
		if err != nil {
			return nil, err
		}
	}

	return f, nil
}

func processSheet(f *excelize.File, sheetName string, data map[string]interface{}) error {
	//获取sheet的所有行
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return err
	}

	//遍历所有的行
	for rowIndex, row := range rows {
		//遍历行中的所有单元格
		for colIndex, cellValue := range row {
			//获取单元格的位置
			cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			//处理单元格
			err = processCell(f, sheetName,cellName ,cellValue, data)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

func processCell(f *excelize.File, sheetName,cellName ,cellValue string, data map[string]interface{}) error {
	//获取模板的值
	repalcedStr, replaced := ReplaceString(cellValue,data)
	if !replaced {
		return nil
	}

	//设置单元格的值
	err := f.SetCellValue(sheetName, cellName, repalcedStr)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

func ReplaceString(orgStr string,data map[string]interface{})(string,bool){
	//识别出过滤参数中的
	re := regexp.MustCompile(`%{([A-Z|a-z|_|0-9|.]*)}`)
	replaceItems:=re.FindAllStringSubmatch(orgStr,-1)
	fmt.Println("ReplaceString replaceItems",replaceItems)
	replaced:=false
	if replaceItems!=nil {
		for _,replaceItem:=range replaceItems {
			repalceStr:=getReplaceString(replaceItem[1],data)
			orgStr=strings.Replace(orgStr,replaceItem[0],repalceStr,-1)
		}
		replaced=true
	}
	return orgStr,replaced
}

func getReplaceString(path string,data map[string]interface{})(string){
	values:=[]string{}
	pathNodes:=strings.Split(path, ".")
	getPathData(pathNodes,0,data,&values)
	//将value转为豆号分割的字符串
	if len(values)>0 {
		valueStr:=strings.Join(values, "\",\"")
		fmt.Println("将value转为豆号分割的字符串","valueStr",valueStr)
		return valueStr
	}
	return path
}

func getPathData(path []string,level int,data map[string]interface{},values *[]string){
	pathNode:=path[level]

	dataNode,ok:=data[pathNode]
	if !ok {
		fmt.Println("getPathData no pathNode ","pathNode",pathNode)
		return
	}

	//如果当前层级为最后一层
	if len(path)==(level+1) {
		switch dataNode.(type) {
			case string:
				sVal, _ := dataNode.(string)   
				*values=append(*values,sVal) 
			case int:
				iVal,_:=dataNode.(int)
				sVal:=fmt.Sprintf("%d",iVal)
				*values=append(*values,sVal)
			case int64:
				iVal,_:=dataNode.(int64)
				sVal:=fmt.Sprintf("%d",iVal)
				*values=append(*values,sVal) 
			default:
				fmt.Println("getPathData not supported value type dataNode type", reflect.TypeOf(dataNode))
		}
	} else {
		//如果不是最后一级，则数据中应该存在list属性
		fmt.Println("getPathData dataNode type is", reflect.TypeOf(dataNode))
		result,ok:=dataNode.(map[string]interface{})
		if !ok {
			fmt.Println("getPathData dataNode is not a map[string]interface{} ")
			return
		}

		//读取list属性
		list,ok:=result["list"]
		if !ok {
			fmt.Println("getPathData dataNode has no List ")
			return
		}

		//list转为数组
		resultList,ok:=list.([]interface{})
		if !ok {
			fmt.Println("getPathData dataNode List is not a []interface{} ")
			return
		}

		for _,row:=range resultList {
			rowMap,ok:=row.(map[string]interface{})
			if !ok {
				fmt.Println("getPathData dataNode row is not a map[string]interface{} ")
				return
			}
			getPathData(path,level+1,rowMap,values)
		}
		return
	}
}