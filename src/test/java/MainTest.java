import org.junit.Test;

public class MainTest {
	
	@SuppressWarnings("static-access")
	@Test
	public void main正常系テスト() throws Exception {
		// パラメータ作成
		String[] args = {"C:\\temp","自由","FUZZY"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
	@SuppressWarnings("static-access")
	@Test
	public void main正常系テスト2() throws Exception {
		// パラメータ作成
		String[] args = {"C:\\temp","あなたですてこうした本意からは背後が進んれるですなしさはしよば得事です。","STRICTLY"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
	@SuppressWarnings("static-access")
	@Test
	public void main正常系テスト3() throws Exception {
		// パラメータ作成
		String[] args = {"C:\\temp","自由","FUZZY","格差"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
	@SuppressWarnings("static-access")
	@Test
	public void main正常系テスト4() throws Exception {
		// パラメータ作成
		String[] args = {"C:\\temp","これはさらに解らないだろ。","STRICTLY","やっていく気持ち"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
}
