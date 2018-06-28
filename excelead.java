package demo1;
import jxl.Workbook;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintStream;

import jxl.Cell;
import jxl.Sheet;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;

import jxl.read.biff.*;

public class excelead {
public String inuputfile;



public void readdata()
{
	File input=new File(inuputfile);
	Workbook w;
	try
	{
		w = Workbook.getWorkbook(input);
		Sheet sheet= w.getSheet(0);
		int i=1;
		while(i<=835)
		{
		Cell cell1 = sheet.getCell(17,i);
        System.out.print(cell1.getContents() + ":");
        Cell cell12 = sheet.getCell(15,i);
        System.out.print(cell12.getContents() + ":");
            // open a connection to the site
           // URL url = new URL("http://argetim2k17.com/ajitesh/sendemail.php");
          //  URLConnection con = url.openConnection();
            // activate the output
          ///  con.setDoOutput(true);
          //  PrintStream ps = new PrintStream(con.getOutputStream());
            // send your parameters to your site
          //  ps.print("&email="+cell1.getContents());
           

            // we have to get the input stream in order to actually send the request
          //  con.getInputStream();

            // close the print stream
          //  ps.close();
            
            //
        if(cell1.getContents()!="")
        {
            URL url=new URL("http://argetim2k17.com/ajitesh/sendemail.php");
            HttpURLConnection httpURLConnection = (HttpURLConnection) url.openConnection();
            httpURLConnection.setRequestMethod("POST");
            httpURLConnection.setDoOutput(true);
            httpURLConnection.setDoInput(true);
            OutputStream outputStream = httpURLConnection.getOutputStream();
            BufferedWriter bufferedWriter = new BufferedWriter(new OutputStreamWriter(outputStream, "UTF-8"));
            String post="";
            post = URLEncoder.encode("email", "UTF-8") + "=" + URLEncoder.encode(cell1.getContents(), "UTF-8")+"&"+
            		URLEncoder.encode("value", "UTF-8") + "=" + URLEncoder.encode(cell12.getContents(), "UTF-8")
                    ;
            bufferedWriter.write(post);
            bufferedWriter.flush();
            bufferedWriter.close();
            InputStream inputStream=httpURLConnection.getInputStream();
            BufferedReader bufferedReader=new BufferedReader(new InputStreamReader(inputStream,"UTF-8"));
            String j;
            StringBuffer sb=new StringBuffer();

            while((j=bufferedReader.readLine())!=null)
            {
                sb.append(j);
            }
            i=i+1;
            System.out.println(sb);
        }
		}   
	}
	catch (IOException e) {
        e.printStackTrace();
    } catch (BiffException e) {
        e.printStackTrace();
    } finally {

        
    }

}

	public static void main(String [] args)
	{
		excelead test=new excelead();
		test.inuputfile="C:\\Users\\kesha\\Desktop\\LIPUN.xls";
		test.readdata();
	}
}