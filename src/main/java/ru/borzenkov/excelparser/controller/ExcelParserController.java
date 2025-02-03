package ru.borzenkov.excelparser.controller;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;
import java.util.PriorityQueue;

@Slf4j
@RestController
@RequestMapping("/api/xlsx")
@RequiredArgsConstructor
@Tag(name = "Чтение XLSX", description = "Сервис для чтения XLSX файлов и поиска N-го максимального значения")
public class ExcelParserController {

    @Operation(summary = "Получение N-го максимального числа из XLSX файла",
            description = "Принимает путь к файлу и число N, возвращает N-е максимальное значение из первого столбца файла")
    @GetMapping("/max-number")
    public Integer getNthMaxNumber(@Parameter(description = "Путь к файлу") @RequestParam String filePath, @Parameter(description = "Номер максимального числа") @RequestParam Integer n) {
        log.info("Запрос на поиск {}-го максимального числа в файле: {}", n, filePath);
        File file = new File(filePath);
        if (!file.exists()) {
            log.error("Файл не найден: {}", filePath);
            throw new IllegalArgumentException("Файл не найден: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            PriorityQueue<Integer> minHeap = new PriorityQueue<>(n + 1);

            for (Row row : sheet) {
                if (Objects.nonNull(row.getCell(0))) {
                    int value = (int) row.getCell(0).getNumericCellValue();
                    minHeap.offer(value);
                    if (minHeap.size() > n) {
                        minHeap.poll();
                    }
                }
            }

            if (minHeap.size() < n) {
                log.error("В файле недостаточно чисел для поиска {}-го максимального значения", n);
                throw new IllegalArgumentException("Недостаточно чисел в файле");
            }
            return minHeap.poll();
        } catch (IOException e) {
            log.error("Ошибка при обработке файла: {}", filePath, e);
            throw new RuntimeException("Ошибка чтения файла", e);
        }
    }
}
