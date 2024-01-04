package myexcel

import (
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelInfo struct {
	Name   string
	Sheets []*SheetInfo
}

type SheetInfo struct {
	Name     string
	Types    []*TypeInfo
	Varnames []string
	Descs    []string
	Content  [][]string
}

func (this *ExcelInfo) Load(path, name string, tag string) error {
	f, err := excelize.OpenFile(path + "/" + name + ".xlsx")
	if err != nil {
		return err
	}
	sheets := f.GetSheetList()
	for _, sheet := range sheets {
		if strings.HasPrefix(sheet, "Sheet") {
			continue
		}
		if len([]byte(sheet)) != len([]rune(sheet)) {
			continue
		}
		sheetInfo := &SheetInfo{}
		sheetInfo.Name = sheet
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		if len(rows) < 4 {
			return errors.New("表结构不足4行:" + name + ".xlsx" + " sheet:" + sheet)
		}
		needExport := make(map[int]bool)
		for i, v := range rows[0] {
			if strings.Contains(v, tag) {
				needExport[i] = true
			}
		}
		for i, vname := range rows[1] {
			if !needExport[i] {
				continue
			}
			sheetInfo.Varnames = append(sheetInfo.Varnames, vname)
		}
		if len(sheetInfo.Varnames) != len(needExport) {
			return errors.New("字段名不能为空:" + name + ".xlsx" + " sheet:" + sheet)
		}
		for i, t := range rows[2] {
			if !needExport[i] {
				continue
			}
			typeInfo, err := getTypeInfoByStr(strings.ToLower(t))
			if err != nil {
				return err
			}
			typeInfo.FixType()
			sheetInfo.Types = append(sheetInfo.Types, typeInfo)
		}
		if len(sheetInfo.Types) != len(needExport) {
			return errors.New("类型不能为空:" + name + ".xlsx" + " sheet:" + sheet)
		}
		for i, desc := range rows[3] {
			if !needExport[i] {
				continue
			}
			sheetInfo.Descs = append(sheetInfo.Descs, desc)
		}
		for i := len(sheetInfo.Descs); i < len(needExport); i++ {
			sheetInfo.Descs = append(sheetInfo.Descs, "")
		}
		for r := 4; r < len(rows); r++ {
			row := rows[r]
			content := []string{}
			for i, cell := range row {
				if !needExport[i] {
					continue
				}
				content = append(content, cell)
			}
			for i := len(content); i < len(needExport); i++ {
				content = append(content, "")
			}
			sheetInfo.Content = append(sheetInfo.Content, content)
		}
		this.Sheets = append(this.Sheets, sheetInfo)
	}
	return nil
}

func (this *ExcelInfo) GenJson(path string) error {
	jsonStr, err := this.ToJson()
	if err != nil {
		return err
	}
	filePath := path + "/" + this.Name + ".json"
	file, err := os.OpenFile(filePath, os.O_RDWR|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if nil != err {
		return errors.New("open json file failed " + this.Name + ".json")
	}
	defer file.Close()
	file.WriteString(jsonStr)
	return nil
}

func (this *ExcelInfo) ToJson() (string, error) {
	ret := "{"
	for i, s := range this.Sheets {
		str, err := s.ToJson()
		if err != nil {
			return "", err
		}
		if str == "" {
			continue
		}
		if i != 0 {
			ret += ","
		}
		ret += "\n"
		ret += str
	}
	ret += "\n}"
	return ret, nil
}

func (this *SheetInfo) ToJson() (string, error) {
	if len(this.Varnames) == 0 {
		return "", nil
	}
	ret := "    \"" + this.Name + "\":["
	repeatCheck := make(map[string]bool)
	needCheck := false
	if this.Name == "Global" {
		needCheck = true
	}
	if this.Varnames[0] == "ID" {
		if this.Types[0].CType != CellTypeSimple || this.Types[0].ValueType1 != "int32" {
			return "", errors.New("ID type err in sheet " + this.Name)
		}
		needCheck = true
	}
	for i, row := range this.Content {
		rowStr := "{"
		for j, cell := range row {
			cellStr := "\"" + this.Varnames[j] + "\":"
			vStr, err := this.Types[j].ParseToJson(cell)
			if err != nil {
				return "", err
			}
			if needCheck && j == 0 {
				if repeatCheck[vStr] {
					return "", errors.New(this.Varnames[0] + " repeat in sheet " + this.Name)
				}
				repeatCheck[vStr] = true
			}
			cellStr += vStr
			if j != 0 {
				rowStr += ","
			}
			rowStr += cellStr
		}
		rowStr += "}"
		if i != 0 {
			ret += ","
		}
		ret += "\n        " + rowStr
	}
	ret += "\n    ]"
	return ret, nil
}

func (this *ExcelInfo) GenCSharpCode(path string) error {
	ret := ""
	ret += "using Newtonsoft.Json;\n"
	ret += "using System.Collections.Generic;\n"
	ret += "using UnityEngine;\n"
	ret += "using YooAsset;\n\n"
	ret += "namespace GameData\n"
	ret += "{\n"
	needMap := make(map[string]bool)
	for _, s := range this.Sheets {
		if len(s.Varnames) == 0 {
			continue
		}
		if s.Varnames[0] == "ID" && s.Types[0].CType == CellTypeSimple && s.Types[0].ValueType1 == "int32" {
			needMap[s.Name] = true
		}
		ret += "    public class " + s.Name + "Info\n"
		ret += "    {\n"
		for i := range s.Varnames {
			ret += "        /// <summary>\n"
			ret += "        /// " + strings.Replace(s.Descs[i], "\n", " ", -1) + "\n"
			ret += "        /// </summary>\n"
			if s.Types[i].CType == CellTypeSimple {
				ret += "        public " + getCSharpType(s.Types[i].ValueType1) + " " + s.Varnames[i] + ";\n"
			} else if s.Types[i].CType == CellTypeSlc {
				ret += "        public List<" + getCSharpType(s.Types[i].ValueType1) + "> " + s.Varnames[i] + ";\n"
			} else if s.Types[i].CType == CellTypeDoubleSlc {
				ret += "        public  List<List<" + getCSharpType(s.Types[i].ValueType1) + ">> " + s.Varnames[i] + ";\n"
			} else if s.Types[i].CType == CellTypeMap {
				ret += "        public Dictionary<" + getCSharpType(s.Types[i].ValueType1) + "," + getCSharpType(s.Types[i].ValueType2) + "> " + s.Varnames[i] + ";\n"
			}
		}
		ret += "    }\n\n"
	}
	ret += "    public class " + this.Name + "Cfg\n"
	ret += "    {\n"
	ret += "        private static " + this.Name + "Cfg _instance;\n"
	ret += "        public static " + this.Name + "Cfg Get()\n"
	ret += "        {\n"
	ret += "            if (_instance == null)\n"
	ret += "            {\n"
	ret += "                _instance = Create();\n"
	ret += "            }\n"
	ret += "            return _instance;\n"
	ret += "        }\n"
	for _, s := range this.Sheets {
		if len(s.Varnames) == 0 {
			continue
		}
		ret += "        public List<" + s.Name + "Info> " + s.Name + ";\n"
		ret += "        private Dictionary<int, " + s.Name + "Info> _" + s.Name + "Dict;\n"
	}
	ret += "        private static " + this.Name + "Cfg Create()\n"
	ret += "        {\n"
	ret += "            AssetHandle handle = YooAssets.LoadAssetSync<TextAsset>(\"Assets/GameRes/Json/" + this.Name + "\");\n"
	ret += "            TextAsset text = handle.AssetObject as TextAsset;\n"
	ret += "            return JsonConvert.DeserializeObject<" + this.Name + "Cfg>(text.text);\n"
	ret += "        }\n\n"
	ret += "        private void InitDict()\n"
	ret += "        {\n"
	for _, s := range this.Sheets {
		if len(s.Varnames) == 0 {
			continue
		}
		ret += "            _" + s.Name + "Dict = new Dictionary<int, " + s.Name + "Info>();\n"
		ret += "            foreach(" + s.Name + "Info info in " + s.Name + ")\n"
		ret += "            {\n"
		ret += "                _" + s.Name + "Dict.Add(info.ID, info);\n"
		ret += "            }\n"
	}
	ret += "        }\n\n"
	for _, s := range this.Sheets {
		if len(s.Varnames) == 0 {
			continue
		}
		ret += "        public " + s.Name + "Info Get" + s.Name + "ByID(int id)\n"
		ret += "        {\n"
		ret += "            if (_" + s.Name + "Dict == null)\n"
		ret += "            {\n"
		ret += "                InitDict();\n"
		ret += "            }\n"
		ret += "            return _" + s.Name + "Dict[id];\n"
		ret += "        }\n"
	}
	ret += "    }\n"
	ret += "}\n"
	filePath := path + "/" + this.Name + "Cfg.cs"
	file, err := os.OpenFile(filePath, os.O_RDWR|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if nil != err {
		return errors.New("open ts code file failed " + this.Name + "Cfg.cs")
	}
	defer file.Close()
	file.WriteString(ret)
	return nil
}

func getCSharpType(t string) string {
	switch t {
	case "int", "int32":
		return "int"
	case "int64":
		return "long"
	case "float", "float32":
		return "float"
	case "float64":
		return "double"
	case "string":
		return "string"
	case "bool":
		return "bool"
	}
	panic(fmt.Sprintf("type %s unknow", t))
}

func (this *ExcelInfo) GenCSharpGlobalKey(path string) error {
	if this.Name != "Global" {
		return errors.New("this is not Global.xlsx")
	}
	if len(this.Sheets) != 1 || this.Sheets[0].Name != "Global" {
		return errors.New("Global sheet error")
	}
	if this.Sheets[0].Types[0].CType != CellTypeSimple || this.Sheets[0].Types[0].ValueType1 != "string" {
		return errors.New("Global key type error")
	}
	ret := ""
	ret += "namespace GameData\n"
	ret += "{\n"
	ret += "    static class GlobalKey\n"
	ret += "    {\n"
	for _, v := range this.Sheets[0].Content {
		ret += "        public const string " + v[0] + " = \"" + v[0] + "\";\n"
	}
	ret += "    }\n"
	ret += "}\n"
	filePath := path + "/GlobalKey.cs"
	file, err := os.OpenFile(filePath, os.O_RDWR|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if nil != err {
		return errors.New("open code file failed GlobalKey.cs")
	}
	defer file.Close()
	file.WriteString(ret)
	return nil
}
