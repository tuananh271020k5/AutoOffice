#!/usr/bin/env -S node --experimental-strip-types
/**
 * Builds `installer/output/AutoOffice-Setup.exe`.
 *
 * 1. Reads `.env` (if present) and updates `manifest.production.xml` when
 *    `VITE_PUBLIC_URL` is configured.
 * 2. If Inno Setup (`ISCC.exe`) is available, compiles `installer/setup.iss`.
 * 3. Otherwise, uses Windows' built-in .NET C# compiler (`csc.exe`) to build
 *    a standalone native Windows GUI installer `installer/output/AutoOffice-Setup.exe`
 *    with `manifest.production.xml` embedded inside.
 */
import {
  existsSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
  rmSync,
} from "node:fs";
import { resolve, join } from "node:path";
import { spawnSync } from "node:child_process";

function parseEnvFile(filePath: string): Record<string, string> {
  if (!existsSync(filePath)) return {};
  const out: Record<string, string> = {};
  const content = readFileSync(filePath, "utf8");
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq === -1) continue;
    const key = line.slice(0, eq).trim();
    const val = line
      .slice(eq + 1)
      .trim()
      .replace(/^['"]|['"]$/g, "");
    out[key] = val;
  }
  return out;
}

function updateProductionManifestUrl(publicUrl: string): void {
  const manifestPath = resolve("manifest.production.xml");
  if (!existsSync(manifestPath)) return;
  const normalized = publicUrl.replace(/\/+$/, "") + "/";
  let xml = readFileSync(manifestPath, "utf8");
  // Replace existing base URL ending with /assets/ or / in manifest.production.xml
  xml = xml.replace(
    /https?:\/\/[^"'\s]+\/assets\/(icon-\d+\.png)/g,
    `${normalized}assets/$1`,
  );
  xml = xml.replace(
    /(<SourceLocation\s+DefaultValue=")[^"]+("\s*\/>)/g,
    `$1${normalized}$2`,
  );
  xml = xml.replace(
    /(<bt:Url\s+id="Taskpane\.Url"\s+DefaultValue=")[^"]+("\s*\/>)/g,
    `$1${normalized}$2`,
  );
  writeFileSync(manifestPath, xml, "utf8");
  console.log(`✓ Updated manifest.production.xml with URL: ${normalized}`);
}

function findInnoSetup(): string | null {
  const candidates = [
    "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
    "C:\\Program Files\\Inno Setup 6\\ISCC.exe",
  ];
  for (const c of candidates) {
    if (existsSync(c)) return c;
  }
  return null;
}

function findCsc(): string | null {
  const candidates = [
    "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe",
    "C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\csc.exe",
  ];
  for (const c of candidates) {
    if (existsSync(c)) return c;
  }
  return null;
}

function buildWithCsc(cscPath: string, outDir: string): number {
  const tempDir = resolve(outDir, ".tmp-build");
  mkdirSync(tempDir, { recursive: true });

  const manifestProdPath = resolve("manifest.production.xml");
  const embeddedManifestPath = join(tempDir, "manifest.xml");
  writeFileSync(embeddedManifestPath, readFileSync(manifestProdPath));

  const appManifestPath = join(tempDir, "app.manifest");
  writeFileSync(
    appManifestPath,
    `<?xml version="1.0" encoding="utf-8"?>
<assembly manifestVersion="1.0" xmlns="urn:schemas-microsoft-com:asm.v1">
  <assemblyIdentity version="1.0.0.0" name="AutoOffice.Installer.app"/>
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
    <security>
      <requestedPrivileges xmlns="urn:schemas-microsoft-com:asm.v3">
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false" />
      </requestedPrivileges>
    </security>
  </trustInfo>
</assembly>`,
    "utf8",
  );

  const csSourcePath = join(tempDir, "Installer.cs");
  writeFileSync(
    csSourcePath,
    `using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AutoOfficeInstaller
{
    static class Program
    {
        private const string ShareName = "AutoOfficeAddin";
        private const string OwnSharePath = @"C:\\AutoOfficeAddin";
        private const string CatalogGuid = "{B2C3D4E5-F6A7-8901-BCDE-F12345678903}";
        private const string TrustedCatalogsKey = @"Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs";
        private const string InstallerStateKey = @"Software\\AutoOffice\\Installer";

        [STAThread]
        static int Main(string[] args)
        {
            bool silent = false;
            bool uninstall = false;
            foreach (var a in args)
            {
                if (a.Equals("/S", StringComparison.OrdinalIgnoreCase) || a.Equals("/SILENT", StringComparison.OrdinalIgnoreCase) || a.Equals("/VERYSILENT", StringComparison.OrdinalIgnoreCase))
                    silent = true;
                if (a.Equals("/UNINSTALL", StringComparison.OrdinalIgnoreCase) || a.Equals("/U", StringComparison.OrdinalIgnoreCase))
                    uninstall = true;
            }

            if (silent)
            {
                string err;
                bool ok = uninstall ? PerformUninstall(out err) : PerformInstall(out err);
                return ok ? 0 : 1;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
            return 0;
        }

        private static void RunCmd(string fileName, string arguments)
        {
            try
            {
                var psi = new ProcessStartInfo(fileName, arguments)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (var p = Process.Start(psi))
                {
                    if (p != null) p.WaitForExit(10000);
                }
            }
            catch { }
        }

        private static string FindHostCatalogUrl()
        {
            try
            {
                using (var baseKey = Registry.CurrentUser.OpenSubKey(TrustedCatalogsKey, false))
                {
                    if (baseKey == null) return "";
                    foreach (var sub in baseKey.GetSubKeyNames())
                    {
                        if (string.Equals(sub, CatalogGuid, StringComparison.OrdinalIgnoreCase))
                            continue;
                        using (var sk = baseKey.OpenSubKey(sub, false))
                        {
                            if (sk == null) continue;
                            var url = sk.GetValue("Url") as string;
                            if (!string.IsNullOrEmpty(url) && url.StartsWith(@"\\\\"))
                                return url;
                        }
                    }
                }
            }
            catch { }
            return "";
        }

        public static bool PerformInstall(out string errorMessage)
        {
            errorMessage = "";
            try
            {
                RunCmd("taskkill", "/F /IM WINWORD.EXE");

                string hostCatalogUrl = FindHostCatalogUrl();
                bool useHostCatalog = !string.IsNullOrEmpty(hostCatalogUrl);
                string computerName = Environment.GetEnvironmentVariable("COMPUTERNAME");
                if (string.IsNullOrEmpty(computerName)) computerName = "localhost";
                string networkPath = @"\\\\" + computerName + @"\\" + ShareName;

                if (!useHostCatalog)
                {
                    string wefCache = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        @"Microsoft\\Office\\16.0\\Wef"
                    );
                    if (Directory.Exists(wefCache))
                    {
                        try { Directory.Delete(wefCache, true); } catch { }
                    }

                    if (!Directory.Exists(OwnSharePath))
                        Directory.CreateDirectory(OwnSharePath);

                    RunCmd("net", "share " + ShareName + " /delete /y");
                    RunCmd("net", "share " + ShareName + "=\\"" + OwnSharePath + "\\" /GRANT:Everyone,FULL");

                    using (var catKey = Registry.CurrentUser.CreateSubKey(TrustedCatalogsKey + @"\\" + CatalogGuid))
                    {
                        if (catKey != null)
                        {
                            catKey.SetValue("Id", CatalogGuid, RegistryValueKind.String);
                            catKey.SetValue("Url", networkPath, RegistryValueKind.String);
                            catKey.SetValue("Flags", 1, RegistryValueKind.DWord);
                        }
                    }
                }

                string targetFile = useHostCatalog
                    ? Path.Combine(hostCatalogUrl, "autooffice.xml")
                    : Path.Combine(OwnSharePath, "manifest.xml");

                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("manifest.xml"))
                {
                    if (s == null)
                    {
                        errorMessage = "Không tìm thấy tài nguyên manifest.xml bên trong bộ cài.";
                        return false;
                    }
                    using (FileStream fs = new FileStream(targetFile, FileMode.Create, FileAccess.Write))
                    {
                        s.CopyTo(fs);
                    }
                }

                using (var stKey = Registry.CurrentUser.CreateSubKey(InstallerStateKey))
                {
                    if (stKey != null)
                        stKey.SetValue("ManifestPath", targetFile, RegistryValueKind.String);
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        public static bool PerformUninstall(out string errorMessage)
        {
            errorMessage = "";
            try
            {
                RunCmd("taskkill", "/F /IM WINWORD.EXE");
                using (var stKey = Registry.CurrentUser.OpenSubKey(InstallerStateKey, false))
                {
                    if (stKey != null)
                    {
                        var mp = stKey.GetValue("ManifestPath") as string;
                        if (!string.IsNullOrEmpty(mp) && File.Exists(mp))
                        {
                            try { File.Delete(mp); } catch { }
                        }
                    }
                }
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\\AutoOffice", false);
                Registry.CurrentUser.DeleteSubKeyTree(TrustedCatalogsKey + @"\\" + CatalogGuid, false);

                if (Directory.Exists(OwnSharePath))
                {
                    RunCmd("net", "share " + ShareName + " /delete /y");
                    try { Directory.Delete(OwnSharePath, true); } catch { }
                }
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private class InstallerForm : Form
        {
            private Button btnInstall;
            private Button btnUninstall;
            private Label lblStatus;

            public InstallerForm()
            {
                Text = "Cài đặt AutoOffice AI (Word, Excel, PowerPoint)";
                Size = new Size(520, 340);
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                StartPosition = FormStartPosition.CenterScreen;
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);

                var lblTitle = new Label
                {
                    Text = "AutoOffice — Trợ lý AI cho Microsoft Office 365",
                    Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                    Location = new Point(24, 20),
                    Size = new Size(460, 30)
                };

                var lblDesc = new Label
                {
                    Text = "Bộ cài đặt sẽ tự động đăng ký Add-in AutoOffice vào Microsoft Word, Excel và PowerPoint trên máy tính này.\\r\\n\\r\\n" +
                           "Sau khi cài đặt xong:\\r\\n" +
                           "1. Mở Microsoft Word (hoặc Excel / PowerPoint)\\r\\n" +
                           "2. Vào thẻ Home → Add-ins → More Add-ins (hoặc Insert → My Add-ins)\\r\\n" +
                           "3. Chọn tab 'Shared Folder' (Thư mục chia sẻ) → Chọn 'AutoOffice' → bấm Add.",
                    Location = new Point(24, 60),
                    Size = new Size(460, 145)
                };

                lblStatus = new Label
                {
                    Text = "Sẵn sàng cài đặt.",
                    ForeColor = Color.FromArgb(0, 102, 204),
                    Location = new Point(24, 212),
                    Size = new Size(460, 28)
                };

                btnInstall = new Button
                {
                    Text = "Cài đặt ngay",
                    Size = new Size(150, 38),
                    Location = new Point(24, 245),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnInstall.Click += OnInstallClick;

                btnUninstall = new Button
                {
                    Text = "Gỡ cài đặt",
                    Size = new Size(120, 38),
                    Location = new Point(186, 245)
                };
                btnUninstall.Click += OnUninstallClick;

                var btnClose = new Button
                {
                    Text = "Đóng",
                    Size = new Size(100, 38),
                    Location = new Point(380, 245)
                };
                btnClose.Click += (s, e) => Close();

                Controls.Add(lblTitle);
                Controls.Add(lblDesc);
                Controls.Add(lblStatus);
                Controls.Add(btnInstall);
                Controls.Add(btnUninstall);
                Controls.Add(btnClose);
            }

            private void OnInstallClick(object sender, EventArgs e)
            {
                btnInstall.Enabled = false;
                lblStatus.Text = "Đang cài đặt AutoOffice Add-in...";
                Application.DoEvents();

                string err;
                if (PerformInstall(out err))
                {
                    lblStatus.ForeColor = Color.Green;
                    lblStatus.Text = "✓ Cài đặt thành công! Hãy mở Word → Home → Add-ins → Shared Folder.";
                    MessageBox.Show(
                        "Đã cài đặt AutoOffice thành công!\\n\\n" +
                        "Cách mở trong Microsoft Word / Excel / PowerPoint:\\n" +
                        "1. Mở ứng dụng Word, Excel hoặc PowerPoint\\n" +
                        "2. Vào thẻ Home → Add-ins → More Add-ins\\n" +
                        "3. Chọn mục 'Shared Folder' → Chọn 'AutoOffice' → Bấm Add.",
                        "Cài đặt hoàn tất",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    lblStatus.ForeColor = Color.Red;
                    lblStatus.Text = "✗ Lỗi cài đặt: " + err;
                    MessageBox.Show("Lỗi khi cài đặt: " + err, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                btnInstall.Enabled = true;
            }

            private void OnUninstallClick(object sender, EventArgs e)
            {
                string err;
                if (PerformUninstall(out err))
                {
                    lblStatus.ForeColor = Color.Green;
                    lblStatus.Text = "✓ Đã gỡ cài đặt AutoOffice.";
                    MessageBox.Show("Đã gỡ cài đặt AutoOffice khỏi máy tính.", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    lblStatus.ForeColor = Color.Red;
                    lblStatus.Text = "✗ Lỗi gỡ cài đặt: " + err;
                }
            }
        }
    }
}
`,
    "utf8",
  );

  const outExe = resolve(outDir, "AutoOffice-Setup.exe");
  const args = [
    "/nologo",
    "/target:winexe",
    `/out:${outExe}`,
    `/win32manifest:${appManifestPath}`,
    `/resource:${embeddedManifestPath},manifest.xml`,
    "/r:System.dll",
    "/r:System.Drawing.dll",
    "/r:System.Windows.Forms.dll",
    csSourcePath,
  ];

  const res = spawnSync(cscPath, args, { stdio: "inherit" });
  rmSync(tempDir, { recursive: true, force: true });

  if (res.status === 0) {
    console.log(`✓ Built standalone installer: ${outExe}`);
    return 0;
  }
  return res.status ?? 1;
}

function main(): number {
  const envVars = {
    ...parseEnvFile(resolve(".env")),
    ...parseEnvFile(resolve(".env.local")),
    ...process.env,
  };

  const publicUrl = (process.argv[2] || envVars.VITE_PUBLIC_URL || "").trim();
  if (publicUrl) {
    updateProductionManifestUrl(publicUrl);
  }

  const outDir = resolve("installer/output");
  mkdirSync(outDir, { recursive: true });

  const iscc = findInnoSetup();
  if (iscc) {
    console.log(`Building installer with Inno Setup (${iscc})...`);
    const res = spawnSync(iscc, [resolve("installer/setup.iss")], {
      stdio: "inherit",
    });
    return res.status ?? 1;
  }

  const csc = findCsc();
  if (csc) {
    console.log(
      "Inno Setup not found — compiling standalone Windows GUI installer via .NET csc.exe...",
    );
    return buildWithCsc(csc, outDir);
  }

  console.error(
    "✗ Neither Inno Setup (ISCC.exe) nor .NET compiler (csc.exe) was found on this system.",
  );
  return 1;
}

process.exit(main());
