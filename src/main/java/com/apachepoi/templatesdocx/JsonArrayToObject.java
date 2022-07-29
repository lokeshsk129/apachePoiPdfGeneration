package com.apachepoi.templatesdocx;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class JsonArrayToObject {

	Map<String, Object> dataConversion;
	public static String filename = "D:/sampleJson3.json";
	public static JSONObject jsonObject;

	public static JSONObject jsonToMap(String filename) throws JSONException, IOException {
		String content = new String(Files.readAllBytes(Paths.get(filename)));
		return new JSONObject(content);
	}

	public Map<String, Object> jsonTojavaMap(JSONObject jsonObject) throws JSONException {
		Map<String, Object> jsonMap = new HashMap<String, Object>();

		if (jsonObject != JSONObject.NULL) {
			jsonMap = toMap(jsonObject);
		}
		return jsonMap;
	}

	private static Map<String, Object> toMap(JSONObject object) throws JSONException {
		Map<String, Object> map = new HashMap<String, Object>();

		Iterator<String> keysItr = object.keys();
		while (keysItr.hasNext()) {
			String key = keysItr.next();
			Object value = object.get(key);

			if (value instanceof JSONArray) {
				value = toList((JSONArray) value);
			}

			else if (value instanceof JSONObject) {
				value = toMap((JSONObject) value);
			}
			map.put(key, value);

		}
		return map;

	}

	private static Object toList(JSONArray array) throws JSONException {
		List<Object> list = new ArrayList<Object>();

		for (int i = 0; i < array.length(); i++) {
			Object value = array.get(i);
			if (value instanceof JSONArray) {
				value = toList((JSONArray) value);

			}

			else if (value instanceof JSONObject) {
				value = toMap((JSONObject) value);
			}
			list.add(value);
		}
		return list;

	}

	public static void main(String[] args) throws IOException, JSONException {

		BufferedReader reader = new BufferedReader(new FileReader(filename));
		StringBuilder stringBuilder = new StringBuilder();
		String line = null;
		String ls = System.getProperty("line.separator");
		while ((line = reader.readLine()) != null) {
			stringBuilder.append(line);
			stringBuilder.append(ls);
		}
		stringBuilder.deleteCharAt(stringBuilder.length() - 1);
		reader.close();
		String content = stringBuilder.toString();
		JSONArray json = new JSONArray(content);

		for (int i = 0; i < json.length(); i++) {
			try {
				jsonObject = json.getJSONObject(i);
			} catch (JSONException e) {
				e.printStackTrace();
			}

		}

	}

}
