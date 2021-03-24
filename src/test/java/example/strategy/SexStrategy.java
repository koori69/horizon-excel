package example.strategy;

import horizon.excel.strategy.DataStrategy;

import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-19 10:05
 */
public class SexStrategy implements DataStrategy {
    public static Map<String, Boolean> map;

    static {
        map = new HashMap<>();
        map.put("女", true);
        map.put("男", false);
    }

    private String key;
    private boolean value;

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public boolean isValue() {
        return value;
    }

    public void setValue(boolean value) {
        this.value = value;
    }

    @Override
    public Object getValue(String key) {
        return map.get(key);
    }

    @Override
    public boolean verify(Object key) {
        boolean flag = false;
        for (String k : map.keySet()) {
            if (k.equals(key)) {
                flag = true;
            }
        }
        return flag;
    }

    @Override
    public String getStrategy() {
        return map.keySet().stream().collect(Collectors.joining(", "));
    }
}
