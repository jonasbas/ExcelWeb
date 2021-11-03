package web.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class ExcelApplicationTests {

	private final ExcelConverter test = new ExcelConverter();
	
	@Test
	void contextLoads() {
	}

	
	@Test
	void testMath() {
		assertEquals(4, test.doMath(2,2,"+"));
		assertEquals(2.5, test.doMath(0.5,5,"*"));
		assertEquals(1, test.doMath(2,2,"/"));
		assertEquals(-2, test.doMath(0,2,"-"));
		assertEquals(0, test.doMath(2,2,"q"));
	}

}
