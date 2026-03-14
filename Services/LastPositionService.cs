using System.Text.Json;
using StartForm.Models;

namespace StartForm.Services
{
    public class LastPositionService
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        private string GetLastPositionDirectory()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "StartForm",
                "last_positions");

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            return dir;
        }

        public void Save(string profileName, List<LastPosition> positions)
        {
            var fileName = SanitizeFileName(profileName) + "_positions.json";
            var filePath = Path.Combine(GetLastPositionDirectory(), fileName);
            var json = JsonSerializer.Serialize(positions, JsonOptions);

            File.WriteAllText(filePath, json);
        }

        public List<LastPosition> Load(string profileName)
        {
            var fileName = SanitizeFileName(profileName) + "_positions.json";
            var filePath = Path.Combine(GetLastPositionDirectory(), fileName);

            if (!File.Exists(filePath))
                return new List<LastPosition>();

            try
            {
                var json = File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<List<LastPosition>>(json, JsonOptions) ?? new List<LastPosition>();
            }
            catch
            {
                return new List<LastPosition>();
            }
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
