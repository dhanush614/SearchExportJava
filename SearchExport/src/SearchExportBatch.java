import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import javax.security.auth.Subject;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.collection.DocumentSet;
import com.filenet.api.collection.IndependentObjectSet;
import com.filenet.api.collection.RepositoryRowSet;
import com.filenet.api.constants.ClassNames;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.core.Connection;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Domain;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.IndependentObject;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.Properties;
import com.filenet.api.property.Property;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.query.RepositoryRow;
import com.filenet.api.query.SearchSQL;
import com.filenet.api.query.SearchScope;
import com.filenet.api.util.UserContext;
import com.filenet.apiimpl.core.FolderImpl;
import com.ibm.casemgmt.api.CaseType;
import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;

public class SearchExportBatch {
	String uri = "http://ibmbaw:9080/wsi/FNCEWS40MTOM";
	String username = "dadmin";
	String password = "dadmin";
	String TOS = "tos";
	UserContext old = null;
	CaseMgmtContext oldCmc = null;

	/**
	 * 
	 */
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
			String sql;

			PropertyFilter pf = new PropertyFilter();
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_SIZE, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_ELEMENTS, null));
			String folderPath = "/Bulk Case Creation Batch";
			Folder myFolder = Factory.Folder.fetchInstance(targetOS, folderPath, null);
			DocumentSet myLoanDocs = myFolder.get_ContainedDocuments();
			Iterator itr = myLoanDocs.iterator();
			while (itr.hasNext()) {
				Document doc = (Document) itr.next();
				doc.fetchProperties(pf);
				SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
				String docCheckInDate = formatter.format(doc.get_DateCheckedIn());
				String todayDate = formatter.format(new Date());
				if (docCheckInDate.equals(todayDate)) {
					sql = createSearchReport(targetOS, doc);
					System.out.println("Query: " + sql);
					SearchSQL searchSql = new SearchSQL(sql);
					RepositoryRowSet rowSetCase = search.fetchRows(searchSql, null, null, new Boolean(true));
					XSSFWorkbook wb = new XSSFWorkbook();
					XSSFSheet sheet = wb.createSheet("Search Results");
					XSSFRow rowhead = sheet.createRow((short) 0);
					Iterator it = rowSetCase.iterator();
					int i = 0, c = 0;
					if (it.hasNext()) {
						RepositoryRow row = (RepositoryRow) it.next();
						Properties props = row.getProperties();
						Iterator propsIt = props.iterator();
						while (propsIt.hasNext()) {
							Property prop = (Property) propsIt.next();
							rowhead.createCell(i).setCellValue(prop.getPropertyName());
						}
						i++;
					}
					while (it.hasNext()) {
						int z = 0;
						RepositoryRow row = (RepositoryRow) it.next();
						XSSFRow excelRow = sheet.createRow((short) ++c);
						Properties props = row.getProperties();
						Iterator propsIt = props.iterator();
						while (propsIt.hasNext()) {
							Property prop = (Property) propsIt.next();
							excelRow.createCell(z).setCellValue(prop.getObjectValue().toString());
							z++;
						}
					}
					OutputStream fileOut = new FileOutputStream(
							"C:\\Users\\Administrator\\Desktop\\DocHandler\\SearchExport.xlsx");
					System.out.println("Excel File has been created successfully.");
					wb.write(fileOut);
					fileOut.close();
				} else {
					System.out.println("No Templates Available, Please upload template and try again..!!");
				}
			}
		} catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
		} finally {
			if (oldCmc != null) {
				CaseMgmtContext.set(oldCmc);
			}
			if (old != null) {
				UserContext.set(old);
			}
		}

	}

	private String createSearchReport(ObjectStore targetOS, Document doc) {
		// TODO Auto-generated method stub
		String query = null;
		try {
			ContentElementList docContentList = doc.get_ContentElements();
			Iterator iter = docContentList.iterator();
			while (iter.hasNext()) {
				ContentTransfer ct = (ContentTransfer) iter.next();
				InputStream stream = ct.accessContentStream();
				XSSFWorkbook workbook = new XSSFWorkbook(stream);
				XSSFSheet sheet = workbook.getSheetAt(0);
				Iterator<Row> rowIterator = sheet.iterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					query = row.getCell(0).getStringCellValue();
				}
			}

		} catch (Exception e) {
			System.out.println(e);
		}
		return query;
	}

}
