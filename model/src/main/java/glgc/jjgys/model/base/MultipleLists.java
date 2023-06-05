package glgc.jjgys.model.base;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;
import java.util.Map;

@Data
public class MultipleLists {
    private List<Map<String, String>> listql;
    private List<Map<String, String>> listsd;
    private List<Map<String, String>> listsurplus;

    public MultipleLists(List<Map<String, String>> listql, List<Map<String, String>> listsd, List<Map<String, String>> listsurplus) {
        this.listql = listql;
        this.listsd = listsd;
        this.listsurplus = listsurplus;
    }
}
