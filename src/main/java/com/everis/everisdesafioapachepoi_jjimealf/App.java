package com.everis.everisdesafioapachepoi_jjimealf;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Clase principal
 *
 */
public class App {
	/** Libro excel */
	private static HSSFWorkbook xlsFile;

	/**
	 * Metodo principal
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		// Genera el documento excel
		generateExcel();

		// Lee el documento excel creado
		readExcel();
	}

	/**
	 * Metodo para generar el documento excel
	 * 
	 */
	public static void generateExcel() {

		/** Lista de alumnos */
		final List<String> alumnNames = new ArrayList<>();
		alumnNames.add("Juan");
		alumnNames.add("Ana");
		alumnNames.add("Paula");
		alumnNames.add("Alberto");
		alumnNames.add("Juanma");
		alumnNames.add("Pablo");
		alumnNames.add("Paula");
		alumnNames.add("Agustin");

		/** Lista notas */
		final List<Integer> alumnNotes = new ArrayList<>();
		alumnNotes.add(6);
		alumnNotes.add(8);
		alumnNotes.add(5);
		alumnNotes.add(7);
		alumnNotes.add(7);
		alumnNotes.add(10);
		alumnNotes.add(9);
		alumnNotes.add(8);

		try {
			// Creacion del libro
			xlsFile = new HSSFWorkbook();
			// Creacion de la pestaña
			final HSSFSheet sheet = xlsFile.createSheet("1ºESO");
			// Creacion de la primera fila del excel
			final HSSFRow row1 = sheet.createRow(0);
			row1.createCell(0).setCellValue("Nombres");
			row1.createCell(1).setCellValue("Notas");
			// Creacion del resto de filas
			HSSFRow row;
			for (int i = 0; i < alumnNames.size(); i++) {
				row = sheet.createRow(i + 1);
				row.createCell(0).setCellValue(alumnNames.get(i));
				row.createCell(1).setCellValue(alumnNotes.get(i));
			}

			// Escritura en el fichero salida
			final FileOutputStream xlsOutFile = new FileOutputStream("NotasESO.xls");
			xlsFile.write(xlsOutFile);
			// Cierre de flujo
			xlsOutFile.close();

		} catch (IOException e) {
			System.out.println("ERROR------ Escritura del XLS Fallida");
		}

	}

	/**
	 * Metodo para leer el Excel
	 * 
	 */
	public static void readExcel() {

		try {
			// Obtencion del libro
			final FileInputStream xlsInFile = new FileInputStream("NotasESO.xls");
			final POIFSFileSystem inputXls = new POIFSFileSystem(xlsInFile);
			// Obtencion de la primera pestaña
			xlsFile = new HSSFWorkbook(inputXls);
			final HSSFSheet sheet = xlsFile.getSheetAt(0);
			// Obtencion de las filas y columnas
			for (int i = 0; i <= sheet.getLastRowNum(); i++) {
				System.out.print(sheet.getRow(i).getCell(0) + "   ");
				System.out.println(sheet.getRow(i).getCell(1));
			}

		} catch (IOException e) {
			System.out.println("ERROR--------Lectura del XLS Fallida");
		}
	}
}
