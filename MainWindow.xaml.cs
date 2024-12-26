using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;

public class MainWindow
{
    private async Task ExtractAndRunSetup()
    {
        try
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "Office_Manager_Temp");
            string setupPath = Path.Combine(tempPath, "setup.exe");
            string configPath = Path.Combine(tempPath, "configuration.xml");

            // Créer le dossier temporaire
            Directory.CreateDirectory(tempPath);

            // Extraire l'outil de déploiement
            using (Stream resource = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("Office_Manager.Resources.officedeploymenttool_18129-20158.exe"))
            using (FileStream file = File.Create(setupPath))
            {
                await resource.CopyToAsync(file);
            }

            // Extraire et exécuter setup.exe
            Process extractProcess = Process.Start(new ProcessStartInfo
            {
                FileName = setupPath,
                WorkingDirectory = tempPath,
                UseShellExecute = true,
                Verb = "runas"
            });
            await Task.Run(() => extractProcess?.WaitForExit());

            // Écrire le fichier de configuration
            File.WriteAllText(configPath, Properties.Resources.configuration);

            // Exécuter setup.exe avec la configuration
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = Path.Combine(tempPath, "setup.exe"),
                Arguments = $"/download {configPath}",
                UseShellExecute = true,
                Verb = "runas",
                WorkingDirectory = tempPath
            };

            using (Process process = Process.Start(startInfo))
            {
                if (process != null)
                {
                    await Task.Run(() => process.WaitForExit());
                }
            }

            // Nettoyer les fichiers temporaires
            try
            {
                Directory.Delete(tempPath, true);
            }
            catch
            {
                // Ignorer les erreurs de nettoyage
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erreur lors de l'exécution: {ex.Message}");
        }
    }
} 