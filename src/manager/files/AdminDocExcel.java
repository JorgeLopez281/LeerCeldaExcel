package manager.files;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AdminDocExcel {

	String rutaArchivo;
	String hojaExcel;
	String nombreArchivoExcel;
	String celdaExcel;
	String datoExcel;

	Properties prop = new Properties();

	/**
	 * Metodo que permite obtener el valor de las propiedades que se encuentra en el
	 * archivo config.properties.
	 * 
	 * @throws IOException
	 */
	private void getProperties() throws IOException {
		InputStream entrada = new FileInputStream("Config.properties");
		prop.load(entrada);

		try {
			rutaArchivo = prop.getProperty("rutaArchivo");
			hojaExcel = prop.getProperty("hojaArchivoExcel");
			celdaExcel = prop.getProperty("celdaArchivoExcel");
			if (rutaArchivo == null) {
				System.out.println("No se encuentra la propiedad rutaArchivo en el archivo de properties");
			}
		} catch (Exception e) {
			System.out.println(
					"Excepcion controlada al intentar cargar la propiedad rutaArchivo del archivo de properties");
		}

	}

	/**
	 * Funcion que obtiene la informacion de una celda especifica en formato de texto  
	 * de un archivo de Excel, no puede ser el resultado de una funcion.
	 * 
	 * @return
	 * @throws IOException
	 */
	public String objManagerFileExcel() throws IOException {
		getProperties();

		File carpeta = new File(rutaArchivo);
		if (carpeta.exists()) {
			nombreArchivoExcel = carpeta.getName();
			try (InputStream file = new FileInputStream(new File(rutaArchivo))) {

//				// Lee el archivo de Excel
//				XSSFWorkbook excelBook = new XSSFWorkbook(file);
//				// Obtiene la hoja del archivo de Excel
//				XSSFSheet excelPage = excelBook.getSheet(hojaExcel);
				
				Workbook wb = WorkbookFactory.create(file);
				Sheet excelPage = wb.getSheet(hojaExcel);

				// Obtiene la referencia de la celda que debe seleccionar
				CellReference ref = new CellReference(celdaExcel);
				// Obtiene la fila dependiendo la celda que se configura en el archivo config.properties
				Row fila = excelPage.getRow(ref.getRow());
				if (fila != null) {
					// Obtiene la columna dependiendo la celda que se configura en el archivo config.properties
					Cell columna = fila.getCell(ref.getCol());
					// Obtiene la informacion que tiene la celda pero no puede ser el resultado de una formula y debe estar en formato texto
					datoExcel = columna.getRichStringCellValue().getString();
					System.out.println("La informacion es: " + datoExcel);
				}
			} catch (Exception e) {
				System.out.println("Error Controlado");
				e.getMessage();
			}
		} else {
			System.out.println("No se encuentra ningun archivo en la ruta especificada");
		}
		// Retorna el dato que se encuentra en la celda.
		return datoExcel;
	}
}
