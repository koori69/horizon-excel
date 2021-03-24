package example.strategy;

import horizon.excel.strategy.DataStrategy;

import java.util.HashMap;
import java.util.Map;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-19 14:30
 */
public class SexExportStrategy implements DataStrategy {
    public static Map<Boolean, String> map;

    static {
        map = new HashMap<>();
        map.put(true, "女");
        map.put(false, "男");
    }

    private boolean key;
    private String value;

    public boolean getKey() {
        return key;
    }

    public void setKey(boolean key) {
        this.key = key;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    @Override
    public Object getValue(String key) {
        switch (key) {
            case "true":
                return map.get(true);
            case "false":
                return map.get(false);
            default:
                return null;
        }
    }
}
