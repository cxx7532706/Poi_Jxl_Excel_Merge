import java.io.File;

public class NameChange {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String path = "高管信息\\高管信息1-69\\";
        File file = new File(path);
        File[] array = file.listFiles();
        for(int j=0;j<array.length;j++){
       	  if(array[j].isFile()){
       		  
       		  String name = array[j].getName();
       		  System.out.println(name);
       		  name = name.replace("[", "(");
       		  name = name.replace("]", ")");
       		  System.out.println(name);
       		  File newFile = new File(path + name);
       		  array[j].renameTo(newFile);
       	  	}
       	 
       	 }
	}

}
