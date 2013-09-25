package com.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Demo {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		String modelPath = "e:/model.docx";
		String destDir = "e:/";
		String photo = "e:/1.gif&8&ssss&100&100";
		
		Map<String, String> map = new HashMap<String,String>();
		map.put("name", "张三");
		map.put("age", 23 + "");
		map.put("address", "八里屯");
		map.put("contact", "59595959");
		map.put("imgphoto", photo);
		
		String destFileName = createNewWordForModel(modelPath, map, destDir,"new_doc.docx");
		
		System.out.println("new Docx file :" + destFileName);
	}

	public static String createNewWordForModel(String modelPath,
			Map<String, String> map,String destDir,String destFileName) throws IOException, InvalidFormatException {
		FileInputStream in =  new FileInputStream(modelPath);
		//重写了官方XWPFDocument类添加createPicture方法
		CustomXWPFDocument document = new CustomXWPFDocument(in); 
		
		XWPFWordExtractor extractor = new XWPFWordExtractor(document); 
		
		//获取docx文档内容
		System.out.println(extractor.getText());
		
		//替换表格中的字段
		Iterator<XWPFTable> it = document.getTablesIterator();
		while(it.hasNext()){
			XWPFTable table = (XWPFTable)it.next();
			int rcount = table.getNumberOfRows();
			for(int i =0 ;i < rcount;i++){
				XWPFTableRow row = table.getRow(i);
				List<XWPFTableCell> cells =  row.getTableCells();
				for (XWPFTableCell cell : cells){
					for (String key : map.keySet()) {
						if(cell.getText().equals("${" + key + "}") && key.startsWith("img")){
							Map<String,String> imgData = getImgData(map.get(key), "&");
							File imgFile = new File(imgData.get("path"));
							
							if(imgFile.exists() && imgFile.isFile()){
								cell.removeParagraph(0);  
								XWPFParagraph pargraph = cell.addParagraph();
								
								document.addPictureData(new FileInputStream(imgFile), Integer.parseInt(imgData.get("type")));  
								document.createPicture(document.getAllPictures().size() - 1, 
										Integer.parseInt(imgData.get("width")), 
										Integer.parseInt(imgData.get("height"))
										,pargraph); 
							}
						}else if (cell.getText().equals("${" + key + "}")){
							cell.removeParagraph(0);
							cell.setText(map.get(key));
						}
					}
				}
			}
		}
		
		//替换段落中的字段
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		for (XWPFParagraph paragraph : paragraphs){
			List<XWPFRun> runs = paragraph.getRuns();
			for (XWPFRun run : runs) {
				String txt = run.getText(0);
				for (String key : map.keySet()) {
					if(txt != null && (txt = txt.trim()).length() > 0 
							&& txt.equals("${" + key + "}")){
						run.setText(map.get(key),0);
					}
				}
			}
		}

        String newFilePath = null;
        if(destFileName != null){
        	newFilePath = destDir + File.separatorChar + destFileName;
        }else{
        	newFilePath = destDir + File.separatorChar + System.currentTimeMillis() + ".docx";
        }
        
        FileOutputStream out = null;
        try{
        	out = new FileOutputStream(newFilePath);
        	document.write(out);
        }finally{
        	if(out != null){
        		out.close();
        	}
        }
		
		return newFilePath;
	}
	
	private static Map<String,String> getImgData(String imgDataStr,String splitStr){
		Map<String,String> imgDataMap = new HashMap<String,String>();
		
		String[] imgDataArr = imgDataStr.split(splitStr != null ? splitStr : "&");
		
		String imgPath = imgDataArr[0];
		String imgType = imgDataArr.length >=2 ? imgDataArr[1] : Document.PICTURE_TYPE_PNG + "";
		String imgName = imgDataArr.length >=3 ? imgDataArr[2] : "img";
		String width = imgDataArr.length >=4 ? imgDataArr[3] : "100";
		String height = imgDataArr.length >=5 ? imgDataArr[3] : "100";
		
		imgDataMap.put("path", imgPath);
		imgDataMap.put("type", imgType);
		imgDataMap.put("name", imgName);
		imgDataMap.put("width", width);
		imgDataMap.put("height", height);
		
		return imgDataMap;
	}
}
