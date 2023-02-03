package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type Person struct {
	UserId   int    `db:"user_id"`
	Username string `db:"username"`
	Sex      string `db:"sex"`
	Email    string `db:"email"`
}

func main() {
	// ListDir()
	// CreateExcel()
	info, _ := fileName(".xlsx")
	if info != "-1" {
		OpenExcel(info)

	}
}

// 获取当前目录下的文件名
func fileName(fileType string) (fileName string, err error) {
	str, _ := os.Getwd()

	infos, err := ioutil.ReadDir(str) //读取全部文件
	if err != nil {
		return "读取失败", err
	}
	fileName = "-1"
	//遍历文件名
	for _, info := range infos {
		comma := strings.LastIndex(info.Name(), ".") //获取最后一个点的位置

		if fileType == info.Name()[comma:] { //通过后缀名判断是不是要找的同类型文件

			comma1 := strings.LastIndex(info.Name(), "(教材目录)")
			if comma1 != -1 { //通过文件最后有没有字符“(教材目录)”，判断是不是导出单子
				// fmt.Println(info.Name()) //获取文件的名称
				fileName = info.Name()
			}
		}

	}
	// info = "找到" + info + "表格文件"
	return fileName, err
}

// 打开excel工作簿
func OpenExcel(info string) {
	//打开一个excel工作簿
	f, err := excelize.OpenFile(info)
	if err != nil {
		fmt.Println(err)
		return
	}
	name := f.GetSheetName(2)
	fmt.Println(name)
	//获取指定表
	rows := f.GetRows(name)
	for _, row := range rows { //遍历所有行
		for _, colCell := range row { //遍历一行
			fmt.Println(colCell, "\t")
		}
	}
}

// 创建excel工作簿
func CreateExcel() {
	//创建一个excel工作簿
	f := excelize.NewFile()
	//创建一个工作表
	sheetNum := f.NewSheet("Sheet2")
	//设置单元格的值
	f.SetCellValue("Sheet2", "A2", "hello excel!")
	f.SetCellValue("Sheet1", "B2", "hello B2!")
	//获取工作表中指定单元格的值
	cell := f.GetCellValue("Sheet1", "B2")
	fmt.Println(cell)
	//设置工作簿的默认工作表
	f.SetActiveSheet(sheetNum)
	//根据指定路径保存文件
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

// 获取全部文件名
func ListDir() (s string, err error) {
	str, _ := os.Getwd()
	// fmt.Println("当前程序路径: ", str) //当前路径

	infos, err := ioutil.ReadDir(str) //读取全部文件
	if err != nil {
		return "读取失败", err
	}
	hostName, _ := os.Hostname() //获取主机名
	// fmt.Println(hostName)
	fileName := "./" + hostName + "这里是全部文件名.txt" //文件名称
	fileName1 := hostName + "这里是全部文件名"           //文件名称
	dstFile, err := os.Create(fileName)          //创建文件
	if err != nil {
		fmt.Println(err.Error())
		return "创建文件失败", err
	}
	defer dstFile.Close()

	// names := make([]string, len(infos))
	for _, info := range infos {
		// fmt.Println(info.Name()) //获取每个文件的名称
		comma := strings.LastIndex(info.Name(), ".") //获取最后一个点的位置

		// fmt.Println(info.Name()[:comma])//获取每个文件 后缀名 点之前的 字符串
		if info.Name()[:comma] != fileName1 && info.Name()[:comma] != "批量获取文件名" && info.Name()[:2] != "~$" {
			fmt.Println(info.Name()[:comma], fileName1)
			dstFile.WriteString(info.Name()[:comma] + "\n") //把文件名写入txt
		}
	}
	info := "写入文档" + fileName1 + "成功!"
	fmt.Println(info)
	return info, err
}
