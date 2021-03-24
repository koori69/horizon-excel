package horizon.excel.strategy;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-19 10:03
 */
public interface DataStrategy {
    default Object getValue(String key) {
        return null;
    }

    default boolean verify(Object key) {
        return true;
    }

    default String getStrategy() {
        return "";
    }
}
