using System.Text.Json;
using StartForm.Models;

namespace StartForm.Services
{
    public class ProfileService
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public string GetProfileDirectory()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "StartForm",
                "profiles");

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            return dir;
        }

        public void Save(Profile profile)
        {
            profile.UpdatedAt = DateTime.Now;

            var fileName = SanitizeFileName(profile.ProfileName) + ".json";
            var filePath = Path.Combine(GetProfileDirectory(), fileName);
            var json = JsonSerializer.Serialize(profile, JsonOptions);

            File.WriteAllText(filePath, json);
        }

        public Profile? Load(string profileName)
        {
            var fileName = SanitizeFileName(profileName) + ".json";
            var filePath = Path.Combine(GetProfileDirectory(), fileName);

            if (!File.Exists(filePath))
                return null;

            var json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<Profile>(json, JsonOptions);
        }

        public List<Profile> LoadAll()
        {
            var dir = GetProfileDirectory();
            var profiles = new List<Profile>();

            foreach (var filePath in Directory.GetFiles(dir, "*.json"))
            {
                try
                {
                    var json = File.ReadAllText(filePath);
                    var profile = JsonSerializer.Deserialize<Profile>(json, JsonOptions);
                    if (profile != null)
                        profiles.Add(profile);
                }
                catch
                {
                    // 破損したJSONファイルはスキップ
                }
            }

            return profiles;
        }

        public void Delete(string profileName)
        {
            var fileName = SanitizeFileName(profileName) + ".json";
            var filePath = Path.Combine(GetProfileDirectory(), fileName);

            if (File.Exists(filePath))
                File.Delete(filePath);
        }

        private static string SanitizeFileName(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (var c in invalidChars)
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
