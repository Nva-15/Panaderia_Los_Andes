package com.proyecto.panaderialosandes.services;
import com.proyecto.panaderialosandes.models.Productos;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import org.springframework.stereotype.Service;
@Service
public class ExcelExportService {
    // escribe el excel directamente en la respuesta para que el navegador lo descargue
    public void generarReporteProductos(List<Productos> productos, OutputStream salida) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Productos");

            // Crear encabezados
            Row headerRow = sheet.createRow(0);
            String[] headers = {"ID", "Nombre", "Categoría", "Precio", "Cantidad", "Unidad", "Estado"};
            CellStyle headerStyle = getHeaderCellStyle(workbook);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // Llenar datos de productos
            int rowNum = 1;
            for (Productos producto : productos) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(producto.getId());
                row.createCell(1).setCellValue(producto.getNombre());
                row.createCell(2).setCellValue(producto.getCategoria_id() != null ? producto.getCategoria_id().getNombre() : "");
                row.createCell(3).setCellValue(producto.getPrecio());
                row.createCell(4).setCellValue(producto.getCantidad());
                row.createCell(5).setCellValue(producto.getUnidad());
                row.createCell(6).setCellValue(producto.getEstado());
            }

            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(salida);
        }
    }

    private CellStyle getHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
}
