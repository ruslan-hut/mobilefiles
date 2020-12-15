import org.jawin.DispatchPtr;
import org.jawin.win32.Ole32;

import java.io.FileReader;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Scanner;

public class FileService {

    private static DispatchPtr app;

    public static void main(String args[]){
        start77();
    }

    private static void start77(){

        final String settingsFile = "mobilefiles.settings";
        try {
            FileReader fileReader = new FileReader(settingsFile);
            Scanner scanner = new Scanner(fileReader);
            while (scanner.hasNextLine()){
                String line = scanner.nextLine();
                if (line.contains("path=")){

                    String init = line.replace("path=", "");

                    if (openConnection(init)){
                        try {
                            app.invoke("MobileFiles");
                        }catch (Exception ex) {
                            errorLog("init string: " + init);
                            errorLog("Invoke MobileFiles: " + convert(ex.toString()));
                        }
                        closeConnection();
                    }
                }
            }
        }catch (Exception fileReaderException){
            errorLog("FileReader: "+fileReaderException.toString());
        }

    }

    private static boolean openConnection(String init){

        boolean classCreated = false;

        try {
            Ole32.CoInitialize();
            app = new DispatchPtr("V77.Application");
            classCreated = true;
        }catch (Exception ex){
            errorLog("Class V77.Application; "+ convert(ex.toString()));
        }

        if (classCreated){
            try {
                app.invoke("Initialize",app.get("RMTrade"),init,"NO_SPLASH_SHOW");
                return true;
            }catch (Exception ex){
                errorLog("Initialize: " + init);
                errorLog("openConnection: "+ convert(ex.toString()));
            }
        }

        closeConnection();
        return false;
    }

    private static void closeConnection(){
        app = null;
        try {
            Ole32.CoUninitialize();
        }catch (Exception ex2){
            errorLog("CoUninitialize: "+ex2.toString());
        }
    }

    private static String convert(String message){
        String result;
        try {
            result = new String(message.getBytes("ISO_8859_1"), "Windows-1251");
        }catch (Exception stringEx){
            result = message;
        }
        return result;
    }

    private static void errorLog(String message){
        final String errorsFile = "mobilefiles.errors";
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        message = simpleDateFormat.format(calendar.getTime())+" "+message+"\n";
        try {
            FileWriter fileWriter = new FileWriter(errorsFile, true);
            fileWriter.write(message);
            fileWriter.flush();
        }catch (Exception ex){
            System.out.println(ex.toString());
        }
    }
}
