package web.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class ExcelController {
	
	@GetMapping("/")
	public String home() {
		return "Main.html";
	}
	
	@PostMapping("/upload")
	public String upload(@RequestParam("uploadFile") MultipartFile file){
		try {
			if(file.isEmpty()) {
				return "NoFile.html";
			}
			else {
				ExcelConverter conv = new ExcelConverter(file);
				Workbook resultWb = conv.excel();
				OutputStream fileOut = new FileOutputStream("tmp.xlsx");
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
	public ResponseEntity<Resource> download() throws FileNotFoundException {
	    String filename = "ergebnisse.xlsx";
	    
	    InputStream in = new FileInputStream("tmp.xlsx");
	    InputStreamResource file = new InputStreamResource(in);
	    
	    return ResponseEntity.ok()
	        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
	        .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
	        .body(file);
	  }
	
}