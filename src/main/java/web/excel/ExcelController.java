package web.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class ExcelController {
	
	//Get Request für die Startseite
	@GetMapping("/")
	public String home() {
		//löschen Datei falls sie nicht gedownloadet wurde
		File file = new File("ergebnisse.xlsx");
		if(file.exists()) {
			file.delete();
		}
		return "Main.html";
	}
	
	//Post Request für den Berechnen Knopf wertet die Excel Datei aus und schreibt eine neue
	@PostMapping("/upload")
	public String upload(@RequestParam("uploadFile") MultipartFile file){
		try {
			//falls keine Datei übergeben wurde 
			if(file.isEmpty()) {
				return "NoFile.html";
			}
			else {
				//übergen die Datei an eine Helferklasse die die Verwertung übernimmt
				ExcelConverter conv = new ExcelConverter(file);
				Workbook resultWb = conv.excel();
				OutputStream fileOut = new FileOutputStream("ergebnisse.xlsx");
				resultWb.write(fileOut);
				fileOut.close();
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return "Download.html";
	}
	
	@GetMapping("/download")
	public void dowloadAndDelete(HttpServletResponse response) throws IOException {
		
	    File file = new File("ergebnisse.xlsx");
	    
	    //Setzen Response Header für den File Download
	    response.setContentType("application/vnd.ms-excel");
	    response.setHeader("Content-disposition", "attachment; filename=" + file.getName());

	    OutputStream out = response.getOutputStream();
	    FileInputStream in = new FileInputStream(file);

	    // in nach out kopieren 
	    IOUtils.copy(in,out);

	    //Schließen Input und Outputstream und löschen die Excel Datei von der Festplatte
	    out.close();
	    in.close();
	    file.delete();
	}

//	Andere Lösung ohne Datei löschen
//	@GetMapping("/download")
//	public ResponseEntity<Resource> download() throws FileNotFoundException {
//	    String filename = "ergebnisse.xlsx";
//	    
//	    InputStream in = new FileInputStream("tmp.xlsx");
//	    InputStreamResource file = new InputStreamResource(in);
//	   
//	    return ResponseEntity.ok()
//	        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
//	        .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
//	        .body(file);
//	  }
	
}