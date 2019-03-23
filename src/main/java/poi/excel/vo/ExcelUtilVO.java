package poi.excel.vo;

/**
 * Excel 공통 Library VO
 * 
 * @author LJW
 */
public class ExcelUtilVO {

    private String		sourceFieldName;		//Excel sheet dataset의 header 명
		private String		targetFieldName;		//VO class의 변수 명
		private Class<?> 	targetFieldType;		//VO class의 변수 형
		private short 		index = -1;					//값의 column index (Excel)

		/**
		 * Getter - Excel sheet dataset의 header 명
		 * 
		 * @author LJW
		 * @return Excel sheet dataset의 header 명
		 */
		public String getSourceFieldName() {
			return this.sourceFieldName;
		}

		/**
		 * Setter - Excel sheet dataset의 header 명
		 * 
		 * @author LJW
		 * @param  sourceFieldName Excel sheet dataset의 header 명
		 */
		public void setSourceFieldName(String sourceFieldName) {
			this.sourceFieldName = sourceFieldName;
		}

		/**
		 * Getter - VO class의 변수 명
		 * 
		 * @author LJW
		 * @return VO class의 변수 명
		 */
		public String getTargetFieldName() {
			return this.targetFieldName;
		}

		/**
		 * Setter - VO class의 변수 명
		 * 
		 * @author LJW
		 * @param  targetFieldName VO class의 변수 명
		 */
		public void setTargetFieldName(String targetFieldName) {
			this.targetFieldName = targetFieldName;
		}
		
		/**
		 * Getter - 값의 column index (Excel)
		 * 
		 * @author LJW
		 * @return 값의 column index (Excel)
		 */
		public short getIndex() {
			return this.index;
		}

		/**
		 * Setter - 값의 column index (Excel)
		 * 
		 * @author LJW
		 * @param  index 값의 column index (Excel)
		 */
		public void setIndex(short index) {
			this.index = index;
		}

		/**
		 * Getter - VO class의 변수 형
		 * 
		 * @author LJW
		 * @return VO class의 변수 형
		 */
		public Class<?> getTargetFieldType() {
			return this.targetFieldType;
		}
	
		/**
		 * Setter - VO class의 변수 형
		 * 
		 * @author LJW
		 * @param  targetFieldType VO class의 변수 형
		 */
		public void setTargetFieldType(Class<?> targetFieldType) {
			this.	targetFieldType = targetFieldType;
		}

}