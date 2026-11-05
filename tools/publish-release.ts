import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

function getGitHubToken(): string {
  const res = spawnSync('git', ['credential', 'fill'], {
    input: 'protocol=https\nhost=github.com\n\n',
    encoding: 'utf8',
  });
  const lines = (res.stdout || '').split(/\r?\n/);
  for (const line of lines) {
    if (line.startsWith('password=')) {
      return line.slice('password='.length).trim();
    }
  }
  throw new Error('Could not retrieve GitHub token from git credential manager');
}

async function main() {
  const token = getGitHubToken();
  const owner = 'anhkkk5';
  const repo = 'AutoOffice';
  const tag = 'v1.1.0';
  const exePath = resolve('installer/output/AutoOffice-Setup.exe');
  const exeBytes = readFileSync(exePath);

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/vnd.github+json',
    'User-Agent': 'AutoOffice-Release-Publisher',
    'X-GitHub-Api-Version': '2022-11-28',
  };

  // Check if release v1.1.0 already exists
  let release: any = null;
  const checkRes = await fetch(`https://api.github.com/repos/${owner}/${repo}/releases/tags/${tag}`, {
    headers,
  });
  if (checkRes.ok) {
    release = await checkRes.json();
  } else {
    const createRes = await fetch(`https://api.github.com/repos/${owner}/${repo}/releases`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tag_name: tag,
        target_commitish: 'master',
        name: 'AutoOffice v1.1.0 — Giao diện tiếng Việt thân thiện (ẩn thuật ngữ code) & 10 Công cụ MCP',
        body: [
          '## 📥 Tải về bộ cài đặt AutoOffice v1.1.0 (Bản mới nhất)',
          '',
          '### ✨ Điểm mới trong v1.1.0:',
          '- Ẩn toàn bộ khối mã kỹ thuật/JSON và thuật ngữ lập trình trong khung chat, chỉ hiển thị trạng thái tiếng Việt ngắn gọn và kết quả trả lời tự nhiên.',
          '- Việt hóa toàn bộ danh mục 10 nhóm tính năng và câu lệnh mẫu.',
          '- Giao diện Responsive tự động co giãn mượt mà khi thu hẹp khung chat bên phải Word.',
          '',
          '### 🛠️ Hướng dẫn cài đặt nhanh (30 giây):',
          '1. Tải và mở **`AutoOffice-Setup.exe`** ở mục **Assets** bên dưới -> Bấm **"Cài đặt AutoOffice"**.',
          '2. Mở **Microsoft Word** -> Vào **Home -> Add-ins -> More Add-ins -> Shared Folder (Thư mục dùng chung)** -> Chọn **AutoOffice** -> Bấm **Add**.',
          '3. Nhập **API Key cá nhân** của bạn vào khung chat AutoOffice để bắt đầu sử dụng!',
        ].join('\n'),
        draft: false,
        prerelease: false,
      }),
    });
    if (!createRes.ok) {
      const errText = await createRes.text();
      throw new Error(`Failed to create release (${createRes.status}): ${errText}`);
    }
    release = await createRes.json();
  }

  // Delete existing asset if already present
  if (Array.isArray(release.assets)) {
    for (const asset of release.assets) {
      if (asset.name === 'AutoOffice-Setup.exe') {
        await fetch(`https://api.github.com/repos/${owner}/${repo}/releases/assets/${asset.id}`, {
          method: 'DELETE',
          headers,
        });
      }
    }
  }

  const uploadUrl = String(release.upload_url).replace(/\{\?name,label\}$/, '') + '?name=AutoOffice-Setup.exe';
  const uploadRes = await fetch(uploadUrl, {
    method: 'POST',
    headers: {
      ...headers,
      'Content-Type': 'application/vnd.microsoft.portable-executable',
      'Content-Length': String(exeBytes.length),
    },
    body: exeBytes,
  });

  if (!uploadRes.ok) {
    const errText = await uploadRes.text();
    throw new Error(`Failed to upload asset (${uploadRes.status}): ${errText}`);
  }

  const uploaded = await uploadRes.json();
  console.log(`✓ Release published: ${release.html_url}`);
  console.log(`✓ Asset download URL: ${uploaded.browser_download_url}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

