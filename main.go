package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	var ourFileDir, inFileName string

	// 根据输入参数数量进行不同处理
	switch len(os.Args) {
	case 3:
		ourFileDir = os.Args[1]
		inFileName = os.Args[2]
	case 2:
		homeDir, err := os.UserHomeDir()
		if err != nil {
			log.Fatalf("获取用户主目录失败: %v", err)
		}

		ourFileDir = filepath.Join(homeDir, "workspace", "xlsx2http")
		inFileName = os.Args[1]
	default:
		log.Fatalf("Usage: %s <ourFileDir> <inFileName> 或 %s <inFileName>", os.Args[0], os.Args[0])
	}

	// 记录输入参数
	log.Printf("输入参数解析成功: 输出目录 = %s, 输入文件 = %s", ourFileDir, inFileName)

	// 判断 ourFileDir 是否存在，如果不存在则创建它
	if _, err := os.Stat(ourFileDir); os.IsNotExist(err) {
		log.Printf("目录 %s 不存在，正在创建...", ourFileDir)
		if err := os.MkdirAll(ourFileDir, os.ModePerm); err != nil {
			log.Fatalf("无法创建目录: %v", err)
		}
		log.Printf("目录 %s 创建成功", ourFileDir)
	}

	// 获取当前时间并格式化为字符串
	currentTime := time.Now().Format("200601021504") // 格式为 YYYYMMDDHHMM
	fileName := filepath.Join(ourFileDir, fmt.Sprintf("output_%s.http", currentTime))

	// 创建或打开输出文件
	outputFile, err := os.Create(fileName)
	if err != nil {
		log.Fatalf("无法创建输出文件: %v", err)
	}
	defer func() {
		if err := outputFile.Close(); err != nil {
			log.Printf("关闭输出文件时出错: %v", err)
		}
	}()

	log.Printf("成功创建输出文件: %s", fileName)

	// 打开Excel文件
	f, err := excelize.OpenFile(inFileName)
	if err != nil {
		log.Fatalf("打开Excel文件失败: %v", err)
	}

	log.Printf("成功打开Excel文件: %s", inFileName)

	// 获取第一个工作表名称
	sheetName := f.GetSheetName(0) // 获取第一个工作表的名称
	if sheetName == "" {
		log.Fatalf("无法获取工作表名称")
	}

	log.Printf("使用工作表: %s", sheetName)

	// 读取所有行数据
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("读取行数据失败: %v", err)
	}

	log.Printf("成功读取工作表行数据，共 %d 行", len(rows))

	// 初始化数据结构
	columnsData, columnsDataID, columnsDataCompanyName := processRows(rows)

	// 输出每一列的数据
	for colIndex := range columnsData {
		joinedData := strings.Join(columnsData[colIndex], ", ")
		companyName := columnsDataCompanyName[colIndex][0]
		numEntries := len(columnsData[colIndex])

		for _, id := range columnsDataID[colIndex] {
			_, err := fmt.Fprintf(outputFile, "### ------------------------------------------------\n"+
				"### %v %d \n"+
				"POST http://address:port/xxxx/xxxxx\n"+
				"Content-Type: application/x-www-form-urlencoded\n\n"+
				"content=%s\n&expressNo=\n&id=%s\n\n\n\n",
				companyName, numEntries, joinedData, id)
			if err != nil {
				log.Printf("写入输出文件失败: %v", err)
			}
		}
	}

	log.Printf("成功写入输出文件")
}

func processRows(rows [][]string) ([][]string, [][]string, [][]string) {
	var (
		columnsData            [][]string
		columnsDataID          [][]string
		columnsDataCompanyName [][]string
		idRegex                = regexp.MustCompile(`ID:(\d+)`)
		re                     = regexp.MustCompile(`[\r\n]+`)
	)

	if len(rows) > 1 {
		numCols := len(rows[0])
		columnsData = make([][]string, numCols)
		columnsDataID = make([][]string, numCols)
		columnsDataCompanyName = make([][]string, numCols)

		for rowIndex, row := range rows {
			for colIndex, cell := range row {
				if rowIndex == 0 { // 第一行：公司名称
					replaced := re.ReplaceAllString(cell, "")
					replaced = strings.TrimSpace(replaced)
					columnsDataCompanyName[colIndex] = append(columnsDataCompanyName[colIndex], replaced)
				} else if rowIndex == 1 { // 第二行：ID
					if match := idRegex.FindStringSubmatch(cell); match != nil {
						columnsDataID[colIndex] = append(columnsDataID[colIndex], match[1])
					}
				} else if cell != "" { // 其他行且非空
					columnsData[colIndex] = append(columnsData[colIndex], cell)
				}
			}
		}
	}

	return columnsData, columnsDataID, columnsDataCompanyName
}
