package glgc.jjgys.system.easyexcel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.util.StringUtils;
import glgc.jjgys.system.exception.JjgysException;

import java.lang.reflect.Field;
import java.util.*;

public abstract class ExcelHandler<T> extends AnalysisEventListener<T>
{
    /**
     * 每隔500条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private int BATCH_COUNT = 500;
    List<T> list = new ArrayList<T>();

    private Class<T> clazz;
    public ExcelHandler() {};
    public ExcelHandler(Class<T> clazz) {
        this.clazz = clazz;
    };
    public ExcelHandler(int bATCH_COUNT)
    {
        BATCH_COUNT = bATCH_COUNT;
    }

    /**重写该函数来处理数据*/
    public abstract void handle(List<T> dataList);

    @Override
    public void invoke(T data, AnalysisContext context)
    {
        list.add(data);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT)
        {
            handle(list);
            list.clear(); // 存储完成清理 list
        }
    }


    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        super.invokeHeadMap(headMap, context);
        if (clazz != null){
            try {
                Map<Integer, String> indexNameMap = getIndexNameMap(clazz);
                Set<Integer> keySet = indexNameMap.keySet();
                for (Integer key : keySet) {
                    if (StringUtils.isEmpty(headMap.get(key))){
                        throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
                    }
                    if (!headMap.get(key).equals(indexNameMap.get(key))){
                        throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
                    }
                }
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
            }
        }

    }

    /**
     *获取注解里ExcelProperty的value，用作校验excel
     * @param clazz
     * @return
     * @throws NoSuchFieldException
     */
    @SuppressWarnings("rawtypes")
    public Map<Integer,String> getIndexNameMap(Class clazz) throws NoSuchFieldException {
        Map<Integer,String> result = new HashMap<>();
        Field field;
        Field[] fields=clazz.getDeclaredFields();
        for (int i = 0; i <fields.length ; i++) {
            field=clazz.getDeclaredField(fields[i].getName());
            field.setAccessible(true);
            ExcelProperty excelProperty=field.getAnnotation(ExcelProperty.class);
            if(excelProperty!=null){
                int index = excelProperty.index();
                String[] values = excelProperty.value();
                StringBuilder value = new StringBuilder();
                for (String v : values) {
                    value.append(v);
                }
                result.put(index,value.toString());
            }
        }
        return result;
    }


    @Override
    public void doAfterAllAnalysed(AnalysisContext arg0)
    {
        handle(list);  // 这里也要保存数据，确保最后遗留的数据也存储到数据库
    }


}