package util;

import java.io.InputStream;
import java.net.URL;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.ConcurrentHashMap;

public class ConfigManager {

	private static ConfigManager configManger = new ConfigManager();

	private static final String defaultFileName = "auto";
	
	private volatile static Map<String, Properties> proMap = new ConcurrentHashMap<String, Properties>();

	private ConfigManager() {

	}

	public static ConfigManager instance() {
		return configManger;
	}

	/**
	 * 
	 * @param profileName
	 *            
	 * @param key
	 * 
	 * @return
	 */
	public String getProperty(String profileName, String key) {
		if (key == null || profileName == null) {
			throw new IllegalArgumentException("key is null");
		}
		Properties properties = proMap.get(profileName);

		if (properties == null) {
			synchronized (this) {

				if (properties == null) {

					properties = this.loadProps(profileName);

					if (properties != null) {
						proMap.put(profileName, properties);
					} else {
						return null;
					}
				}
			}
		}

		String value = properties.getProperty(key);
		return value;
	}
	public void clearProps(String proFileName) {
		proMap.remove(proFileName);
	}
	public String getProperty(String key) {
		return getProperty(defaultFileName,key);
	}
	/**
	 * 
	 * @param profileName
	 * @param key
	 * @return
	 */
	public int getIntProperty(String profileName, String key) {
		if (key == null || profileName == null) {
			throw new IllegalArgumentException("key is null");
		}

		String intValue = this.getProperty(profileName, key);

		try {
			return Integer.parseInt(intValue);
		} catch (Exception e) {
			e.printStackTrace();
			return -1;
		}
	}
	public int getIntProperty(String key) {
		return getIntProperty(defaultFileName,key);
	}

	public Properties loadProps(String proFileName) {

		Properties properties = null;

		InputStream in = null;

		try {
			properties = new Properties();
			if (proFileName != null && proFileName.startsWith("http:")) {
				URL url = new URL(proFileName);
				in = url.openStream();
			} else {
				String fileName = "/" + proFileName + ".properties";
				in = getClass().getResourceAsStream(fileName);
				properties.load(in);
			}
			properties.load(in);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (in != null)
					in.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		return properties;
	}
}
