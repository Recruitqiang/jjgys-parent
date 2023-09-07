package glgc.jjgys.system.utils;

import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.util.Zip4jConstants;
import org.apache.commons.compress.utils.IOUtils;
import org.springframework.mock.web.MockMultipartFile;


import org.springframework.web.multipart.MultipartFile;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.List;




public class JjgFbgcUtils {


    /**
     * http响应zip
     * @param srcDir
     * @param out
     * @throws ZipException
     * @throws IOException
     */
    public static void zipFile(String srcDir, OutputStream out ) throws ZipException, IOException {




        // 要打包的文件夹
        File currentFile = new File(srcDir);
        // 生成的压缩文件
        final File file = new File(srcDir + ".zip");

        ZipFile zipFile = new ZipFile(file);

        ZipParameters parameters = new ZipParameters();
        // 压缩方式
        parameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE);
        // 压缩级别
        parameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_FASTEST);


        File[] fs = currentFile.listFiles();
        // 遍历test文件夹下所有的文件、文件夹
        for (File f : fs) {
            if (f.isDirectory()) {
                zipFile.addFolder(f.getPath(), parameters);
            } else {
                zipFile.addFile(f, parameters);
            }
        }
        try( InputStream fis = new FileInputStream(zipFile.getFile())) {
            out.flush();
            IOUtils.copy(fis, out);
            IOUtils.closeQuietly(fis);
            IOUtils.closeQuietly(out);
        }
        //deleteDirAndFiles(file);
       file.delete();


    }

    /**
     * multipartFile转file
     */
    public static File multipartFileToFile(MultipartFile multipartFile) {
        File file = null;
        InputStream inputStream = null;
        OutputStream outputStream = null;
        try {
            inputStream = multipartFile.getInputStream();
            file = new File(multipartFile.getOriginalFilename());
            outputStream = new FileOutputStream(file);
            write(inputStream, outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return file;
    }
    public static void write(InputStream inputStream, OutputStream outputStream) {
        byte[] buffer = new byte[4096];
        try {
            int count = inputStream.read(buffer, 0, buffer.length);
            while (count != -1) {
                outputStream.write(buffer, 0, count);
                count = inputStream.read(buffer, 0, buffer.length);
            }
        } catch (RuntimeException e) {
            throw e;
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }
    /**
     * file 转 multipartFile
     */
    public static MultipartFile getMultipartFile(File file) {
        try {
            MultipartFile cMultiFile = new MockMultipartFile(file.getName(), file.getName(), null, new FileInputStream(file));
            return cMultiFile;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {

        }


    }
    /**
     * @return    suffix  目录名称      pathName  目录根路径
     * @Description //TODO  创建一个临时目录
     **/
    public static java.nio.file.Path createDirectory(String suffix, String pathName) {
        java.nio.file.Path path1 = Paths.get(pathName + "/" + suffix);
        //判断目录是否存在
        boolean pathExists = Files.exists(path1, new LinkOption[]{LinkOption.NOFOLLOW_LINKS});
        if (pathExists) {
            return path1;
        }
        try {
            path1 = Files.createDirectory(path1);

        } catch (IOException e) {
            e.printStackTrace();
        }

        return path1;
    }
    /**
     * @return      OriginalPathName   源文件    path  目标文件地址    pathName  根目录
     * @Description //TODO  复制文件
     **/
    public static void copyFile(String OriginalPathName, java.nio.file.Path path, String pathName) {
        java.nio.file.Path originalPath = Paths.get(pathName + "/" + OriginalPathName);

        try {
            File dbfFile = new File(path.toUri());
            if (!dbfFile.exists() ) {
                dbfFile.createNewFile();
            }
            Files.copy(originalPath, path,
                    StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * @return   attachments  这个参数主要是从数据库中获取文件的信息，这里可以根据自己的需求更改

     * @Description //TODO 向文件夹中添加文件
     **/
    public static void addFile(List<File> files, String path, String pathName) {
        //遍历所有的文件
        int i = 0;
        for (File file : files) {
            i++;
            String name=file.getName();
            String type = name.substring(name.lastIndexOf("."));
            java.nio.file.Path targetPath = Paths.get(path + "/"+name);
            copyFile(file.getName(), targetPath, pathName);
        }
    }




    /**
     * @Description //TODO  删除临时目录及文件
     **/
    public static void deleteDirAndFiles(File file) {
        File[] files = file.listFiles();
        if (files != null && files.length > 0) {
            for (File f : files) {
                deleteDirAndFiles(f);
            }
        }
        file.delete();
        //System.out.println("删除成功" + file.getName());
    }












}
