package main

import (
	"excelParsing/app"
	"fmt"
	"log"
	"sync"
)

func main() {

	var wg sync.WaitGroup
	sliceID, err := app.GetDataFromMainExcel("./mainFile/main.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	mapOfID, err := app.GetUniqueValues(sliceID) //map, которая содержит уникальные номера товаров
	if err != nil {
		log.Fatalln(err)
	}
	fileNames, err := app.GetNamesOfFiles()
	if err != nil {
		log.Fatalln(err)
	}

	parsingCount := len(fileNames)
	if parsingCount == 0 {
		log.Fatalln("ошибка. Файлы не найдены")
	}
	parsingResults := make(chan app.DataFromFile, parsingCount)
	for idWorker, name := range fileNames {
		wg.Add(1)
		fmt.Println("запуск горутины для файла", name)
		id := idWorker + 1
		go func(idRoutine int, name string) {
			defer wg.Done()
			log.Println("запуск горутины id=", idRoutine)
			data, err := app.GetDataFromFiles(name, idRoutine)
			if err != nil {
				log.Printf("Worker %d error: %v", idRoutine, err)
				return
			}
			parsingResults <- data
			log.Printf("Результаты работы парсера %v записаны в канал", idWorker)
		}(id, name)

	}

	wg.Wait()
	close(parsingResults)
	log.Println("Все данные приняты, канал закрыт")

	var fullParsingResult app.DataFromFile
	fullParsingResult, err = app.GetFullData(parsingResults)
	if err != nil {
		log.Fatalln(err)
	}

	chanResults := make(chan app.LineResult, len(mapOfID))

	for word := range mapOfID {
		wg.Add(1)
		go func(wordId string) { //ищем все совпадения для каждого слова. Значения выводим в канал
			defer wg.Done()
			listForId, err := app.GetValuesForItem(wordId, fullParsingResult)
			if err != nil {
				fmt.Println(err)
				return
			}

			chanResults <- app.LineResult{
				Id:   wordId,
				Data: listForId,
				Err:  err,
			}

		}(word)
	}
	wg.Wait()
	close(chanResults)
	log.Println("Поиск значений выполнен. Канал закрыт")

	resultMap, err := app.GetMap(chanResults)
	if err != nil {
		fmt.Println(err)
	}

	err = app.NewFileResult(resultMap)
	if err != nil {
		fmt.Println(err)
	}
}
