import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * エクセルファイルのgrepツール。
 * <pre>
 * 指定したフォルダからエクセルファイルを再帰的に検索し、
 * 指定した文字列をgrepします。
 * 
 * 設定する引数は以下の通り。
 * ・第一引数：検索対象のディレクトリパス
 * ・第二引数：検索対象文字列
 * ・第三引数：検索モード(あいまい検索/完全一致)
 * ・第四引数：置換文字列
 * </pre>
 * 
 * @author TakumiEra
 * 
 */
public class Main {
	
	/**
	 * 検索対象文字列。
	 */
	private static String SEARCH_WORD;
	
	/**
	 * 検索モード。
	 */
	private static SearchMode SEARCH_MODE;
	
	/**
	 * 置換文字列。
	 */
	private static String REPLACE_WORD;
	
	/**
	 * 検索モードのENUM。
	 * 
	 * @author TakumiEra
	 *
	 */
	enum SearchMode{
		/**
		 * あいまい検索。
		 */
		FUZZY,
		/**
		 * 完全一致検索。
		 */
		STRICTLY;
	}
	
	/**
	 * 本処理。
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		// 引数から情報を取得
		String searchDirPath = args[0];
		SEARCH_WORD = args[1];
		SEARCH_MODE = SearchMode.valueOf(args[2]);
		// 置換処理は任意
		if(args.length > 3){
		    REPLACE_WORD = args[3];
		}

		// メッセージを表示
		System.out.println("以下の条件でgrep検索を実行します。");
		System.out.println("検索対象フォルダ：" + searchDirPath);
		System.out.println("検索文字列：" + SEARCH_WORD);
		System.out.println("検索方法：" + SEARCH_MODE);	
		if(args.length > 3){
			System.out.println("置換文字列：" + REPLACE_WORD);
		}
		
		// 検索の実行
		searchDir(searchDirPath);

	}

	/**
	 * 指定したパスを検索する。
	 * 
	 * @param aPath ディレクトリパス
	 * @return
	 * @throws Exception
	 */
	private static boolean searchDir(final String aPath) throws Exception {
		File file = new File(aPath);
		File[] listFiles = file.listFiles(new FilenameFilter() {

			public boolean accept(File aDir, String aName) {
				// ドットで始まるファイルは対象外
				if (aName.startsWith(".")) {
					return false;
				}

				// 対象要素の絶対パスを取得
				String absolutePath = aDir.getAbsolutePath() + File.separator + aName;

				// エクセルファイルのみ対象とする。
				if (new File(absolutePath).isFile()
						&& (absolutePath.endsWith(".xls") || absolutePath
								.endsWith(".xlsx"))) {
					return true;
				} else {
					// ディレクトリの場合、再び同一メソッドを呼出す。
					try {
						return searchDir(absolutePath);
					} catch (Exception e) {
						e.printStackTrace();
						return false;
					}
				}
			}
		});

		if (listFiles == null) {
			return false;
		}

		// 検索を実行
		for (File f : listFiles) {
			if (f.isFile()) {
				searchWord(f);
			}
		}
		System.out.println("検索が完了しました。");
		return true;
	}

	/**
	 * 対象のエクセルシートから文字列を検索し、リストに格納します。
	 * 
	 * @param aFile ファイルオブジェクト
	 * @param SEARCH_WORD 検索対象文字列
	 * @throws Exception
	 */
	private static void searchWord(File aFile) throws Exception {
		// Excelファイルの読込み
		InputStream inputStream = new FileInputStream(aFile);
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(inputStream);
		} catch (Exception ex) {
			return;
		}
		inputStream.close();

		// シート枚数を読込み
		int numberOfSheets = workbook.getNumberOfSheets();
		// シート枚数分ループ処理
		for (int i = 0; i < numberOfSheets; i++) {
			StringBuilder stringBuffer = new StringBuilder();
			// シート
			Sheet sheet = workbook.getSheetAt(i);
			// シート名
			String sheetName = sheet.getSheetName();
			// シートの最終行
			int lastRowNum = sheet.getLastRowNum();
			// 最終行までループ処理
			for (int j = 0; j <= lastRowNum; j++) {
				// 行の取得
				Row row = sheet.getRow(j);
				if (row != null) {
					// 行内の最後のセルの位置
					short lastCellNum = row.getLastCellNum();
					// 行内の最後のセルまでループ処理
					for (int k = 0; k < lastCellNum; k++) {
						// セルを取得
						Cell cell = row.getCell(k);
						if (cell != null) {
							// セルのタイプを取得
							int cellType = cell.getCellType();
							// セルの値が文字列の場合
							if (cellType == Cell.CELL_TYPE_STRING) {
								if(SEARCH_MODE == SearchMode.FUZZY){
									// あいまい検索
								    if (cell.getStringCellValue().contains(SEARCH_WORD)) {
									    stringBuffer = appendRecord(stringBuffer,
													  aFile.getAbsolutePath(),
													  sheetName,
													  convertCellPos(j, k),
													  cell.getStringCellValue());
									    // 置換処理を実行
									    if(REPLACE_WORD != null){
									    	Pattern pattern = Pattern.compile(SEARCH_WORD);
									    	Matcher matcher = pattern.matcher(cell.getStringCellValue());
									    	cell.setCellValue(matcher.replaceAll(REPLACE_WORD));
									    }
								    }
								} else if(SEARCH_MODE == SearchMode.STRICTLY){
									// 完全一致検索
									if (cell.getStringCellValue().equals(SEARCH_WORD)) {
									    stringBuffer = appendRecord(stringBuffer,
													  aFile.getAbsolutePath(),
													  sheetName,
													  convertCellPos(j, k),
													  cell.getStringCellValue());
									    // 置換処理を実行
									    if(REPLACE_WORD != null){
									    	cell.setCellValue(REPLACE_WORD);
									    }
								    }
								}
							}
							// セルの値が数値の場合
							else if (cellType == Cell.CELL_TYPE_NUMERIC) {
								if(SEARCH_MODE == SearchMode.FUZZY){
									// あいまい検索
								    if (String.valueOf(cell.getNumericCellValue()).contains(SEARCH_WORD)) {
									    stringBuffer = appendRecord(stringBuffer,
													  aFile.getAbsolutePath(),
													  sheetName,
													  convertCellPos(j, k),
													  String.valueOf(cell.getNumericCellValue()));
									    // 置換処理を実行
									    if(REPLACE_WORD != null){
									    	Pattern pattern = Pattern.compile(SEARCH_WORD);
									    	Matcher matcher = pattern.matcher(cell.getStringCellValue());
									    	cell.setCellValue(matcher.replaceAll(REPLACE_WORD));
									    }
								    }
								} else if(SEARCH_MODE == SearchMode.STRICTLY){
									// 完全一致検索
									if (String.valueOf(cell.getNumericCellValue()).equals(SEARCH_WORD)) {
									    stringBuffer = appendRecord(stringBuffer,
													  aFile.getAbsolutePath(),
													  sheetName,
													  convertCellPos(j, k),
													  String.valueOf(cell.getNumericCellValue()));
									    // 置換処理を実行
									    if(REPLACE_WORD != null){
									    	cell.setCellValue(REPLACE_WORD);
									    }
								    }
								}
							}
						}
					}
					if (stringBuffer.length() > 0 && !stringBuffer.toString().endsWith("\n")) {
						stringBuffer.append("\n");
					}
				}
			}
			System.out.print(stringBuffer.toString());
		}
		// 上書き保存
		if(REPLACE_WORD != null){
			FileOutputStream outputStream = new FileOutputStream(aFile);
		    workbook.write(outputStream);
		    outputStream.close();
		}
	}
	
    /**
     * セルの位置情報を返す。
     * <pre>
     * 引数の行番号とカラム番号から、セルの位置情報を特定し返却する。
     * 例えば左上のセルは"A1"となる。
     * </pre>
     *     
     * @param aRowNum (０から始まる)行番号
     * @param aColNum (０から始まる)カラム番号
     * @return セルを位置を表す文字列
     * @throws Exception
     */
	private static String convertCellPos(int aRowNum, int aColNum) throws Exception {
		// カラムを表すアルファベットの配列を生成
		final char[] charArray = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
		final int charSize = charArray.length;
		// オフセットを取得
		int offset = aColNum / charSize;

		String cellPos = "";
		if (offset == 0) {
			cellPos = String.valueOf(charArray[aColNum]);
		}
		else if (offset < charSize) {
			cellPos = String.valueOf(charArray[offset - 1])
						+ String.valueOf(charArray[aColNum - charSize * offset]);
		}
		else {
			throw new Exception("範囲外のセルが指定されています。");
		}
		return String.format("%s%d", cellPos, aRowNum + 1);
	}
	
    /**
     * 文字列バッファにファイルパス、シート名、セルの位置情報、値を設定して返却する。
     *
     * @param aStringBuilder 文字列バッファ
     * @param aFilePath ファイルパス
     * @param aSheetName シート名
     * @param aPosition セルの位置情報
     * @param aValue セルの値
     * @return 文字列バッファ
     */
	private static StringBuilder appendRecord(StringBuilder aStringBuilder,
											  String aFilePath,
											  String aSheetName,
											  String aPosition,
											  String aValue) {
		aStringBuilder.append(aFilePath);
		aStringBuilder.append("\t");
		aStringBuilder.append(aSheetName);
		aStringBuilder.append("\t");
		aStringBuilder.append(aPosition);
		aStringBuilder.append("\t");
		aStringBuilder.append(aValue);

		return aStringBuilder;
	}
}
