import java.util.List;

/**
 * 表信息
 *
 * @author: 杜云章
 * @Date: 2020-05-27 16:05
 */
public class TableInfo {

    /**
     * 表名
     */
    private String code;

    /**
     * 驼峰
     */
    private String camel;

    /**
     * 描述
     */
    private String name;

    private List<FieldInfo> fieldInfos;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getCamel() {
        return camel;
    }

    public void setCamel(String camel) {
        this.camel = camel;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<FieldInfo> getFieldInfos() {
        return fieldInfos;
    }

    public void setFieldInfos(List<FieldInfo> fieldInfos) {
        this.fieldInfos = fieldInfos;
    }
}
