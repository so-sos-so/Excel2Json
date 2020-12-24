using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using UnityEngine;

public static class ConfigManager
{
	public static Dictionary<int, Troops> Troops = new Dictionary<int, Troops>();
	public static Dictionary<int, Level> Level = new Dictionary<int, Level>();

    private static string JsonPath = $"{Application.streamingAssetsPath}/Table/";

    public static void Load()
    {
		Troops = LoadOneConfig<Troops>("Troops");
		Level = LoadOneConfig<Level>("Level");

    }

    private static Dictionary<int, T> LoadOneConfig<T>(string configName)
    {
        var path = Path.Combine(JsonPath, configName);
        var jsonContent = ResLoad(path);
        return JsonConvert.DeserializeObject<Dictionary<int, T>>(jsonContent);
    }

    private static string ResLoad(string path)
    {
        var jsonTxt = Resources.Load<TextAsset>(path);
        Debug.Assert(jsonTxt != null, $"{path} can not find");
        return jsonTxt.text;
    }

	public static Troops GetTroopsById(int id)
	{
		if(Troops.TryGetValue(id, out var data))
			return data;
		return null;
	}
	public static Level GetLevelById(int id)
	{
		if(Level.TryGetValue(id, out var data))
			return data;
		return null;
	}

}