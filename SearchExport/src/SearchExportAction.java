import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Iterator;
import javax.security.auth.Subject;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import com.filenet.api.collection.IndependentObjectSet;
import com.filenet.api.constants.ClassNames;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.core.Connection;
import com.filenet.api.core.Domain;
import com.filenet.api.core.Factory;
import com.filenet.api.core.IndependentObject;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.Property;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.query.SearchSQL;
import com.filenet.api.query.SearchScope;
import com.filenet.api.util.UserContext;
import com.filenet.apiimpl.core.FolderImpl;

import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;

public class SearchExportAction {
	String uri = "http://ibmbaw:9080/wsi/FNCEWS40MTOM";
	String username = "dadmin";
	String password = "dadmin";
	String TOS = "tos";
	UserContext old = null;
	CaseMgmtContext oldCmc = null;

	public void getSearchResults() {
		try {
			Connection conn = Factory.Connection.getConnection(uri);
			Subject subject = UserContext.createSubject(conn, username, password, "FileNetP8WSI");
			UserContext.get().pushSubject(subject);

			Domain domain = Factory.Domain.fetchInstance(conn, null, null);
			System.out.println("Domain: " + domain.get_Name());
			System.out.println("Connection to Content Platform Engine successful");
			ObjectStore targetOS = (ObjectStore) domain.fetchObject(ClassNames.OBJECT_STORE, TOS, null);
			System.out.println("Object Store =" + targetOS.get_DisplayName());

			SimpleVWSessionCache vwSessCache = new SimpleVWSessionCache();
			CaseMgmtContext cmc = new CaseMgmtContext(vwSessCache, new SimpleP8ConnectionCache());
			oldCmc = CaseMgmtContext.set(cmc);

			SearchScope search = new SearchScope(targetOS);
			String casePropertyValues[] = new String[] { "LA_Name", "LA_DOB", "LA_Address" };
			String casePropertyValues1 = "LA_Name, LA_DOB, LA_Address, classDescription";
			String caseTypeValue = "LA_LoanProcessingCaseType";
			String searchCriteria = "LA_Address='MAD'";
			String sql = "SELECT casePropertyValues FROM caseTypeValue WHERE searchCriteria";
			sql = sql.replaceAll("casePropertyValues", casePropertyValues1);
			sql = sql.replaceAll("caseTypeValue", caseTypeValue);
			sql = sql.replaceAll("searchCriteria", searchCriteria);
			System.out.println("Query: " + sql);
			SearchSQL searchSQL = new SearchSQL(sql);
			IndependentObjectSet independentObjectSet = search.fetchObjects(searchSQL, new Integer(500), null,
					new Boolean(true));
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("Search Results");
			HSSFRow rowhead = sheet.createRow((short) 0);
			for (int i = 0; i < casePropertyValues.length; i++) {
				rowhead.createCell(i).setCellValue(casePropertyValues[i]);
			}
			Iterator p = null;
			if (!(independentObjectSet.isEmpty())) {
				Iterator it = independentObjectSet.iterator();
				int c = 0;
				while (it.hasNext()) {
					FolderImpl folder = (FolderImpl) it.next();
					HSSFRow row = sheet.createRow((short) ++c);
					for (int i = 0; i < casePropertyValues.length; i++) {
						row.createCell(i)
								.setCellValue(folder.getProperty(casePropertyValues[i]).getObjectValue().toString());
					}
					System.out.println("Row Num: " + c);
				}
			}
			OutputStream fileOut = new FileOutputStream(
					"C:\\Users\\Administrator\\Desktop\\DocHandler\\SearchExport.xlsx");
			System.out.println("Excel File has been created successfully.");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			System.out.println(e);
		} finally {
			if (oldCmc != null) {
				CaseMgmtContext.set(oldCmc);
			}
			if (old != null) {
				UserContext.set(old);
			}
		}

	}

}
