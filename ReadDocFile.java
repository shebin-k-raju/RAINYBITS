package experion;
import java.io.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

public class ReadDocFile
{
    public static void main(String[] args)
    {	
    	BufferedReader br=new BufferedReader(new InputStreamReader(System.in));
        File file = null;
        WordExtractor extractor = null;
        try
        {
        	System.out.println("Enter the File directory:");
        	String dir=br.readLine().replace("\\","\\\\");
            file = new File(dir);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            HWPFDocument document = new HWPFDocument(fis);
            extractor = new WordExtractor(document);
            String[] fileData = extractor.getParagraphText();
            String Name="Not Found",Email="Not Found",Address="Not Found",Mob="Not Found";
            for (int i = 0; i <fileData.length; i++)
            {	int ind=0;
                String temp=fileData[i].toLowerCase();
                if(Name=="Not Found" && temp.contains("name")) {
                	ind=temp.indexOf("name");
                	temp=temp.substring(ind+4).replace(":"," ");
                	temp=temp.replace("  "," ");
                	Name=temp;
                	//System.out.println(temp);
                }
                if(Address=="Not Found" && temp.contains("address")) {
                	ind=temp.indexOf("address");
                	temp=temp.substring(ind+8).replace(":"," ");
                	temp=temp.replace("  "," ");
                	Address=temp;
                	//System.out.println(temp);
                }
                if(Email=="Not Found" && temp.contains("email")) {
                	ind=temp.indexOf("mailto");
                	temp=temp.substring(ind+6);
                	temp=temp.substring(0,temp.indexOf("\""));
                	temp=temp.replace("  "," ");
                	Email=temp;
                	//System.out.println(temp);
                }
                if(Mob=="Not Found" && (temp.contains("mob")||temp.contains("phone")||temp.contains("contact no"))) {
                	ind=temp.indexOf("phone");
                	temp=temp.substring(temp.indexOf(":"));
                	temp=temp.replace("  "," ");
                	Mob=temp;
                	//System.out.println(temp);
                }
            }
            System.out.println("Name:"+Name+"\nEmail:"+Email+"\nMobile:"+Mob+"\nAddress:"+Address);
        }
        catch (Exception exep)
        {
            exep.printStackTrace();
        }
    }
}