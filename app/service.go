package app

import (
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type Line struct {
	Id, Fullname, Consignee, Netto, Brutto, Count, Change, Country string
}

// DataFromFiles содержит набор строк со значениями
type DataFromFile struct {
	Line []Line
}

// MapWithData - содержит данные о товарах, загруженных из файлов
type MapWithData struct {
	IdProduct map[string]DataFromFile
}

type LineResult struct {
	Id   string
	Data DataFromFile
	Err  error
}

// GetDataFromMainExcel - открывает основной файл, который будет находиться в каталоге ./mainFile . Возвращает слайс с ID товаров
func GetDataFromMainExcel(name string) ([]string, error) {
	f, err := excelize.OpenFile(name)
	if err != nil {
		return nil, fmt.Errorf("ошибка при открытии файла %v", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	log.Printf("Файл %s успешно открыт", name)
	rows, err := f.GetRows("Заявка на участие в процедуре")
	if err != nil {
		return nil, fmt.Errorf("ошибка получения строк: %v", err)
	}

	// Собираем данные из 3-го столбца (индекс 2, так как индексация с 0)
	var result []string
	for _, row := range rows {
		if len(row) > 8 && len(row[2]) == 12 { // Проверяем, что в строке есть хотя бы 3 столбца, делаем выборку с номерами ID
			result = append(result, row[2]) // Добавляем значение 3-го столбца
		}
	}

	dataCount := len(result) //количество вхождений

	log.Printf("Чтение файла %s завершено успешно. Количество прочитанных позиций: %v", name, dataCount)

	return result, nil
}

// GetUniqueValues возвращает слайс уникальных номеров
func GetUniqueValues(incoming []string) (map[string]struct{}, error) {
	if incoming == nil {
		return nil, fmt.Errorf("input slice is nil")
	}
	valueMap := make(map[string]struct{})
	for _, val := range incoming {
		if _, exist := valueMap[val]; !exist {
			valueMap[val] = struct{}{}

		}
	}
	log.Println("Количество уникальных записей:", len(valueMap))
	return valueMap, nil
}

// GetNamesOfFiles - сканирует папку и получает имена файлов с расшинением .xlsx, возвращает слайс string
func GetNamesOfFiles() ([]string, error) {
	log.Println("Сканирование папки на наличие файлов для чтения")
	namesList := make([]string, 0)
	files, err := os.ReadDir("./filesExcel/")
	if err != nil {
		return nil, fmt.Errorf("ошибка при получении списка файлов: %v", err)
	}
	for _, file := range files {
		if !file.IsDir() && strings.HasSuffix(file.Name(), ".xlsx") {
			fileName := "./filesExcel/" + file.Name()
			namesList = append(namesList, fileName)
		}
	}
	if len(namesList) == 0 || namesList == nil {
		return nil, fmt.Errorf("файлов для чтения не обнаружено")
	}
	log.Println("Количество файлов готовых для использования:", len(namesList))
	return namesList, nil
}

// GetDataFromFiles возвращает структуру DataFromFiles для каждого файла
func GetDataFromFiles(name string, idWorker int) (DataFromFile, error) {
	var resultData DataFromFile
	log.Printf("Worker: %v. начал работу с файлом %s", idWorker, name)
	f, err := excelize.OpenFile(name)
	if err != nil {
		return resultData, fmt.Errorf("worker: %v ошибка при открытии файла %v", idWorker, err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	log.Printf("Worker: %v.Файл %s успешно открыт", idWorker, name)
	rows, err := f.GetRows("Заявка на участие в процедуре")
	if err != nil {
		return resultData, fmt.Errorf("worker: %v.ошибка получения строк: %v", idWorker, err)
	}

	for _, row := range rows {
		if len(row) > 8 && len(row[2]) == 12 && row[8] != "" { // Проверяем, что в строке есть хотя бы 3 столбца, делаем выборку с номерами ID
			Line := Line{
				Id:        row[2],
				Fullname:  row[3],
				Consignee: row[5],
				Netto:     row[8],
				Brutto:    row[9],
				Count:     row[12],
				Change:    row[15],
				Country:   row[18],
			}
			resultData.Line = append(resultData.Line, Line)
		}
	}
	dataCount := len(resultData.Line) //количество вхождений

	log.Printf("Worker: %v. Чтение файла %s завершено успешно. Количество прочитанных позиций: %v", idWorker, name, dataCount)

	return resultData, nil
}

// GetfullData Объединяет данные из канала в одну структуру
func GetFullData(data <-chan DataFromFile) (DataFromFile, error) {
	resultData := DataFromFile{}
	resultData.Line = make([]Line, 0)
	for d := range data {
		if len(d.Line) == 0 {
			continue
		}
		resultData.Line = append(resultData.Line, d.Line...)

	}
	if len(resultData.Line) == 0 {
		return resultData, fmt.Errorf("no data received from channel")
	}
	log.Printf("Чтение всех файлов завершено. Количество записей: %v", len(resultData.Line))
	return resultData, nil
}

// GetValuesForItem - ищет наличие товара по ID во всех загруженных позициях
func GetValuesForItem(id string, list DataFromFile) (DataFromFile, error) {
	if id == "" {
		return DataFromFile{}, fmt.Errorf("ID слова не может быть пустым")
	}
	var result DataFromFile
	for _, val := range list.Line {
		if id != val.Id {
			continue
		}
		result.Line = append(result.Line, val)
	}
	if len(result.Line) == 0 {
		return DataFromFile{}, fmt.Errorf("предупреждение: нет данных для id %s", id)
	}
	return result, nil
}

// GetMap формирует итоговый список товаров со значениями, которые удалось достать из файлов
func GetMap(data <-chan LineResult) (MapWithData, error) {

	resultMap := MapWithData{
		IdProduct: make(map[string]DataFromFile),
	}
	for value := range data {
		if value.Err != nil {
			return resultMap, value.Err
		}
		resultMap.IdProduct[value.Id] = value.Data
	}
	log.Println("Сборка map со всеми значениями завершена")
	return resultMap, nil

}

func NewFileResult(data MapWithData) error {
	if err := os.MkdirAll("./work", 0755); err != nil {
		return fmt.Errorf("не удалось создать папку work: %v", err)
	}
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			log.Println("Ошибка закрытия файла:", err)
		}
	}()

	// 3. Используем лист по умолчанию "Sheet1"
	sheetName := "Sheet1"

	// 4. Записываем заголовки
	headers := []string{
		"ID", "Наименование", "Грузополучатель",
		"Вес нетто", "Вес брутто", "Цена без НДС",
		"Замена", "Страна",
	}

	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue(sheetName, cell, header)
	}

	// 5. Записываем данные из MapWithData
	row := 2 // Начинаем со второй строки
	for productID, dataFromFile := range data.IdProduct {
		for _, line := range dataFromFile.Line {
			// Записываем все поля структуры Line
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), productID)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), line.Fullname)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), line.Consignee)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), line.Netto)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), line.Brutto)
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), line.Count)
			f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), line.Change)
			f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), line.Country)

			row++
		}
	}

	colWidths := []float64{20, 80, 20, 5, 5, 12, 80, 15}
	for col, width := range colWidths {
		colName, _ := excelize.ColumnNumberToName(col + 1)
		f.SetColWidth(sheetName, colName, colName, width)
	}

	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DDEBF7"}, Pattern: 1},
	})
	f.SetCellStyle(sheetName, "A1", "H1", headerStyle)

	fileName := "./work/products_report.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		log.Fatal("Ошибка сохранения файла:", err)
	}

	log.Printf("Отчет успешно создан: %s\n", fileName)
	log.Printf("Всего товаров: %d\n", len(data.IdProduct))
	return nil
}

// CopyMainFile создает рабочую копию основного файла
func CopyMainFile(filePath string) (string, error) {
	// Создаем папку для рабочих файлов
	workDir := "./work"
	if err := os.MkdirAll(workDir, 0755); err != nil {
		return "", fmt.Errorf("ошибка создания рабочей папки: %v", err)
	}

	// Формируем имя копии с timestamp
	ext := filepath.Ext(filePath)
	filename := strings.TrimSuffix(filepath.Base(filePath), ext)
	timestamp := time.Now().Format("20060102_150405")
	copyName := fmt.Sprintf("%s_work_%s%s", filename, timestamp, ext)
	copyPath := filepath.Join(workDir, copyName)

	// Копируем файл
	src, err := os.Open(filePath)
	if err != nil {
		return "", fmt.Errorf("ошибка открытия исходного файла: %v", err)
	}
	defer src.Close()

	dst, err := os.Create(copyPath)
	if err != nil {
		return "", fmt.Errorf("ошибка создания копии: %v", err)
	}
	defer dst.Close()

	if _, err := io.Copy(dst, src); err != nil {
		return "", fmt.Errorf("ошибка копирования данных: %v", err)
	}

	log.Printf("Создана рабочая копия: %s", copyPath)
	return copyPath, nil
}

// PullDataInCopy заполняет файл-копию
func PullDataInCopy(filePath string, data MapWithData) error {
	if filePath == "" {
		return fmt.Errorf("путь к файлу пустой")
	}
	log.Println("Запущена процедура записи данных в файл", filePath)
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("ошибка открытия файла: %v", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Println("Ошибка закрытия файла:", err)
		}
	}()

	sheetName := "Заявка на участие в процедуре"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("ошибка чтения листа: %v", err)
	}

	updatedCount := 0
	startRow := 17 // Таблица начинается с 17-й строки

	for rowIdx := startRow - 1; rowIdx < len(rows); rowIdx++ {
		row := rows[rowIdx]
		if len(row) < 3 { // Проверяем, что строка содержит как минимум 3 колонки
			continue
		}

		productID := strings.TrimSpace(row[2]) // ID товара в 3-й колонке (C)
		if dataFromFile, exists := data.IdProduct[productID]; exists && len(dataFromFile.Line) > 0 {
			line := dataFromFile.Line[0] // Берем первую (и единственную) запись
			excelRowNum := rowIdx + 1    // Нумерация строк в Excel начинается с 1

			// Обновляем значения в соответствующих колонках
			f.SetCellValue(sheetName, fmt.Sprintf("I%d", excelRowNum), line.Netto)
			f.SetCellValue(sheetName, fmt.Sprintf("J%d", excelRowNum), line.Brutto)
			f.SetCellValue(sheetName, fmt.Sprintf("M%d", excelRowNum), line.Count)
			f.SetCellValue(sheetName, fmt.Sprintf("P%d", excelRowNum), line.Change)
			f.SetCellValue(sheetName, fmt.Sprintf("S%d", excelRowNum), line.Country)
			f.SetCellValue(sheetName, fmt.Sprintf("T%d", excelRowNum), line.Consignee)

			updatedCount++
		}
	}

	// Добавляем заголовок "Исходный пункт назначения" в T16
	f.SetCellValue(sheetName, "T16", "Исходный пункт назначения")

	// Сохраняем изменения
	if err := f.Save(); err != nil {
		return fmt.Errorf("ошибка сохранения файла: %v", err)
	}

	log.Printf("Файл успешно обновлен. Обновлено строк: %d", updatedCount)
	return nil
}

// GetMaxValues - отбирает значение с максимальной стоимостью для каждого id
func GetMaxValues(data MapWithData) (MapWithData, error) {
	result := MapWithData{
		IdProduct: make(map[string]DataFromFile),
	}

	log.Println("Запущен процесс отбора данных с максимальной стоимостью")
	for productID, dataFromFile := range data.IdProduct {
		if len(dataFromFile.Line) == 0 {
			continue
		}

		maxLine := dataFromFile.Line[0]
		maxCount, _ := strconv.Atoi(dataFromFile.Line[0].Count)

		for _, line := range dataFromFile.Line[1:] {
			currentCount, err := strconv.Atoi(line.Count)
			if err != nil {
				continue
			}

			if currentCount > maxCount {
				maxCount = currentCount
				maxLine = line
			}
		}

		// Добавляем в результат
		result.IdProduct[productID] = DataFromFile{
			Line: []Line{maxLine},
		}
	}

	log.Println("Закончен процесс отбора данных с максимальной стоимостью. Количество элементов с данными:", len(result.IdProduct))

	return result, nil
}
