package com.jcg.csv2excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.commons.collections4.map.HashedMap;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.FillPatternType;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

import com.mongodb.BasicDBObjectBuilder;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;

public class ErrorFieldColor {
	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet1 = workbook.createSheet("Sheet1");

		MongoClient client = new MongoClient("localhost", 27017);

		MongoDatabase database = client.getDatabase("admin");

		MongoCollection<Document> collection = database.getCollection("EmpErrorDtls");

		DB db = client.getDB("admin");
		DBCollection coll = db.getCollection("EmployeeDtls");

		FindIterable<Document> iterDoc = collection.find();

		Iterator it = iterDoc.iterator();
		int rownum = 0;
		int headerCount = 0;
		while (it.hasNext()) {
			Document document = (Document) it.next();

			Map<String, Object> mapEmpDtls = new HashedMap<String, Object>(document);

			Set<String> key = mapEmpDtls.keySet();

			for (String object : key) {

				if (mapEmpDtls.containsKey(object)) {

					Object value = mapEmpDtls.get(object);

					System.out.println("Key : " + object + " value :" + value);

					if (object.equals("EmpID")) {
						DBObject query = BasicDBObjectBuilder.start().add("_id", value).get();
						System.out.println(query);
						DBCursor cursor = coll.find(query);
						for (DBObject obj : cursor) {
							System.out.println(mapEmpDtls + "================");
							Map<String, String> map1 = obj.toMap();
							System.out.println(map1);
							Set<String> headers = map1.keySet();

							Row row;

							int colNum = 0;
							if (rownum == 0) {
								row = sheet1.createRow(rownum);
								for (String header : headers) {
									Cell cell1 = row.createCell(colNum++);

									cell1.setCellValue(header);

									headerCount++;
								}
								rownum++;
							}
							colNum = 0;
							row = sheet1.createRow(rownum);
							int tempcount = 0;
							for (String header : headers) {
								Object val = map1.get(header);
								Cell cell1 = row.createCell(colNum++);
								System.out.println(header);
								cell1.setCellValue(val.toString());
								if (header.equals(mapEmpDtls.get("FieldName"))) {
									System.out
											.println("Header  " + header + "FieldName " + mapEmpDtls.get("FieldName"));

									XSSFCellStyle style1 = workbook.createCellStyle();

									style1.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
									style1.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
									style1.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
									style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
									cell1.setCellStyle(style1);

								}
								tempcount++;
								if (tempcount > headerCount) {
									Row tempRow = sheet1.getRow(0);
									Cell tempCell = tempRow.createCell(tempcount - 1);
									System.out.println(header);
									tempCell.setCellValue(header);
									headerCount = tempcount;
								}
							}
							rownum++;

						}
					}
				}
			}

		}
		FileOutputStream outputStream = new FileOutputStream("EmployeeErrorData.xlsx");
		workbook.write(outputStream);
		System.out.println("DOne");
	}

}
