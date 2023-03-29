package glgc.jjgys.system.exception;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class JjgysException extends RuntimeException{

    private Integer code;
    private String msg;

}
