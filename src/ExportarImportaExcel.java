import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.ListIterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportarImportaExcel {
	
	//Explicación exportar e importar-> http://www.javasobretodo.es/programacion/excel-importarexportar-con-java/
	
	public static boolean exportExcel(String nombreHoja, Map<String, Object[]> data, String fileName) {

		// Creamos el libro de trabajo
		XSSFWorkbook libro = new XSSFWorkbook();

		// Creacion de Hoja
		XSSFSheet hoja = libro.createSheet(nombreHoja);

		// Iteramos el map e insertamos los datos
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			// cramos la fila
			Row row = hoja.createRow(rownum++);
			// obtenemos los datos de la fila
			Object[] objArr = data.get(key);
			int cellnum = 0;
			// iteramos cada dato de la fila
			for (Object obj : objArr) {
				// Creamos la celda
				Cell cell = row.createCell(cellnum++);
				// Setteamos el valor con el tipo de dato correspondiente
				if (obj instanceof String)
				cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
				cell.setCellValue((Integer) obj);
			}
		}
		try {
			// Escribimos en fichero
			FileOutputStream out = new FileOutputStream(new File(fileName));
			libro.write(out);
			//cerramos el fichero y el libro
			out.close();
			libro.close();
			System.out.println("Excel exportado correctamente\n");
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	
	}
	
	public static ArrayList<String[]> importExcel(String fileName, int numColums) {

		// ArrayList donde guardaremos todos los datos del excel
		ArrayList<String[]> data = new ArrayList<>();

		try {
			// Acceso al fichero xlsx
			FileInputStream file = new FileInputStream(new File(fileName));

			// Creamos la referencia al libro del directorio dado
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Obtenemos la primera hoja
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterador de filas
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// Iterador de celdas
				Iterator<Cell> cellIterator = row.cellIterator();
				// contador para el array donde guardamos los datos de cada fila
				int contador = 0;
				// Array para guardar los datos de cada fila
				// y añadirlo al ArrayList
				String[] fila = new String[numColums];
				// iteramos las celdas de la fila
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					// Guardamos los datos de la celda segun su tipo
					switch (cell.getCellType()) {
					// si es numerico 
					case Cell.CELL_TYPE_NUMERIC:
						fila[contador] = (int) cell.getNumericCellValue() + "";
						break;
					// si es cadena de texto
					case Cell.CELL_TYPE_STRING:
						fila[contador] = cell.getStringCellValue() + "";
						break;
					}
					// Si hemos terminado con la ultima celda de la fila
					if ((contador + 1) % numColums == 0) {
						// Añadimos la fila al ArrayList con todos los datos
						data.add(fila);
					}
					// Incrementamos el contador
					// con cada fila terminada al redeclarar arriba el contador,
					// no obtenemos excepciones de ArrayIndexOfBounds
					contador++;
				}
			}
			// Cerramos el fichero y workbook
			file.close();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("Excel importado correctamente\n");

		return data;
	}
	
	
	public static void main(String[] args) {

		//Datos a escribir en map(Object[])
	    Map<String, Object[]> data = new TreeMap<String, Object[]>();
	    //Cabecera
	    data.put("1", new Object[] {"ID", "PC", "Nombre", "Apellidos"});
	    //Datos
	    data.put("2", new Object[] {1, 1, "Jesus", "Roldan"});
	    data.put("3", new Object[] {2, 2, "David", "Jimenez"});
	    data.put("4", new Object[] {3, 3, "Iván", "Perez"});
	    data.put("5", new Object[] {4, 4, "Amelia", "Paniagua"});
	    data.put("6", new Object[] {5, 5, "Rafa", "Álvarez"});
	    data.put("7", new Object[] {6, 6, "Antonio", "Garcia"});
	    data.put("8", new Object[] {7, 7, "Jose Antonio", "Rivera"});
	    data.put("9", new Object[] {8, 8, "Pedro", "López"});
	    data.put("10", new Object[] {9, 9, "Manuel", "Martin"});
	    data.put("11", new Object[] {10, 10, "Ramón", "Garcia"});
	    data.put("12", new Object[] {11, 11, "Diego", "Garcia"});
	    data.put("13", new Object[] {12, 12, "Javier", "Garcia"});
	    data.put("14", new Object[] {13, 13, "Luciano", "Garcia"});
	    data.put("15", new Object[] {14, 14, "Jose", "Garcia"});
	    data.put("16", new Object[] {15, 15, "Vicente", "Garcia"});
	    data.put("17", new Object[] {16, 16, "Manuel Jesús", "Garcia"});
	    data.put("18", new Object[] {17, 17, "Francisco", "Garcia"});
	    data.put("19", new Object[] {18, 18, "Laura", "Garcia"});
	    data.put("20", new Object[] {19, 19, "Maria", "Garcia"});
	    data.put("21", new Object[] {20, 20, "Lucia", "Garcia"});
	    data.put("22", new Object[] {21, 21, "Rosa", "Garcia"});
	    
	    boolean correcto = exportExcel("DatosPersonas",data,"Excel.xlsx");
	    
	    if(correcto){
	    	
	    	ArrayList<String[]> datosExcel = importExcel("Excel.xlsx",4);
	        ListIterator<String[]> it = datosExcel.listIterator();
	        
	        while (it.hasNext()) {
				String[] datos =  it.next();
				String personaInfo = "";
				for (String fila : datos) {
					personaInfo += fila + " ";
				}
				System.out.println(personaInfo+"\n");
			}
	    }

	}
		
	

}
