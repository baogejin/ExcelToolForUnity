package main

import (
	"ExcelToolForUnity/myexcel"
	"fmt"
	"os"
	"strings"
)

func main() {
	//生成代码
	rootPath := "../"
	excelPath := rootPath + "Excel"
	jsonPath := rootPath + "Assets/GameRes/Json"
	codePath := rootPath + "Assets/Scripts/Data/GenCode"
	tag := "c" //c为客户端，s为服务器，unity项目默认c

	fileList, err := getExcelList(excelPath)
	if err != nil {
		fmt.Println(err)
		return
	}
	//读取excel
	excels := loadExcels(excelPath, fileList, tag)
	//生成json
	genJson(jsonPath, excels)
	//生成c#代码
	genCSharpCode(codePath, excels, tag)

	fmt.Println("按回车退出")

	fmt.Scanln() // wait for Enter Key
}

func getExcelList(path string) ([]string, error) {
	ret := []string{}
	fileInfoList, err := os.ReadDir(path)
	if err != nil {
		return ret, err
	}
	for _, fileInfo := range fileInfoList {
		name := fileInfo.Name()
		if strings.HasPrefix(name, "~") {
			continue
		}
		strSlc := strings.Split(name, ".")
		if len(strSlc) != 2 || strSlc[1] != "xlsx" {
			continue
		}
		if len([]byte(name)) != len([]rune(name)) {
			fmt.Println("文件名请不要用英文数字以外的字符:", name, "该文件已被忽略")
			continue
		}
		ret = append(ret, strSlc[0])
	}
	return ret, nil
}

func loadExcels(path string, files []string, tag string) []*myexcel.ExcelInfo {
	ret := []*myexcel.ExcelInfo{}
	for i, file := range files {
		excel := &myexcel.ExcelInfo{}
		excel.Name = file
		err := excel.Load(path, file, tag)
		if err != nil {
			panic("load " + file + ".xlsx failed," + err.Error())
		}
		ret = append(ret, excel)
		fmt.Printf("[%d/%d]load %s.xlsx success\n", i+1, len(files), file)
	}
	return ret
}

func genJson(path string, excels []*myexcel.ExcelInfo) {
	for i, excel := range excels {
		if err := excel.GenJson(path); err != nil {
			panic("gen json failed:" + excel.Name + ".xlsx " + err.Error())
		}
		fmt.Printf("[%d/%d]gen json %s.xlsx success\n", i+1, len(excels), excel.Name)
	}
}

func genCSharpCode(path string, excels []*myexcel.ExcelInfo, tag string) {
	if tag == "s" {
		return
	}
	for i, excel := range excels {
		if excel.Name == "Global" {
			if err := excel.GenCSharpGlobalKey(path); err != nil {
				panic("gen global key failed " + err.Error())
			}
			fmt.Printf("[%d/%d]gen global key %s.xlsx success\n", i+1, len(excels), excel.Name)
			continue
		}
		if err := excel.GenCSharpCode(path); err != nil {
			panic("gen code failed:" + excel.Name + ".xlsx " + err.Error())
		}
		fmt.Printf("[%d/%d]gen code %s.xlsx success\n", i+1, len(excels), excel.Name)
	}
}
