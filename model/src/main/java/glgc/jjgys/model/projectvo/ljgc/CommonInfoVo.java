package glgc.jjgys.model.projectvo.ljgc;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class CommonInfoVo implements Serializable {

    private static final long serialVersionUID = 1L;

    private String proname;
    private String htd;
    private String fbgc;
    private String sjz;
    private String sdsjz;
    private String qlsjz;
    private String fhlmsjz;
}
