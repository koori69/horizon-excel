package horizon.excel.formatter;

import horizon.excel.strategy.DataStrategy;

import java.util.HashMap;
import java.util.Map;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 10:07
 */
public class ExcelDataFormatter {
    private Map<String, DataStrategy> formatter = new HashMap();

    public void set(String key, DataStrategy val) {
        formatter.put(key, val);
    }

    public DataStrategy get(String key) {
        return formatter.get(key);
    }
}
