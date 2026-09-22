# Xoá `rsvp_data.db` và `RSVP_tool.pdf` khỏi lịch sử git

Hai file này đã bị commit vào repo public trong commit `48060db` và `9a67145`. Chúng chứa dữ
liệu cá nhân thật: tên, địa chỉ email công ty, số tiền đóng góp, và ảnh chụp màn hình
đánh dấu tài liệu nội bộ. PR #1 đã bỏ theo dõi hai file và thêm chúng vào `.gitignore`,
nhưng việc đó chỉ ngăn các commit **sau này**. Mọi commit cũ vẫn chứa hai file, và ai cũng
tải về được.

Hướng dẫn này là để **chủ repo tự chạy**. Các bước này viết lại toàn bộ lịch sử và
force-push, **không thể hoàn tác**, nên không được chạy tự động.

> Hãy coi dữ liệu này là **đã bị lộ**. Việc xoá khỏi lịch sử giúp không lộ thêm, nhưng
> không thu hồi được bản đã có người clone hoặc fork. Nếu công ty có quy trình báo cáo sự
> cố dữ liệu cá nhân, hãy làm theo quy trình đó.

Mọi lệnh bên dưới chạy trong **PowerShell trên Windows**.

---

## 0. Việc làm ngay, trước khi viết lại lịch sử

1. **Chuyển repo sang private**: GitHub → Settings → General → Danger Zone →
   *Change repository visibility*. Việc này dừng lộ dữ liệu ngay lập tức, và có thể đổi lại
   sau.
2. **Sao lưu hai file** ra ngoài thư mục repo, ví dụ `C:\Backup\rsvp\`. Bạn vẫn cần
   `rsvp_data.db` để chạy app.
3. **Merge hoặc đóng PR #1 trước.** Nếu viết lại lịch sử khi PR còn mở, nhánh của PR cũng
   phải viết lại. Làm một lần, sau khi các nhánh đã ổn định, thì đơn giản hơn.

## 1. Cài `git-filter-repo`

```powershell
py -m pip install git-filter-repo
git filter-repo --version
```

Nếu `git filter-repo` báo không tìm thấy lệnh, có thể thư mục `Scripts` của Python không
nằm trong `PATH`. Khi đó dùng `py -m git_filter_repo` thay cho `git filter-repo` ở mọi bước
bên dưới.

## 2. Viết lại lịch sử trên một bản clone mới

Hãy dùng bản clone mới. **Không dùng thư mục bạn đang làm việc hằng ngày.**

```powershell
cd C:\Temp
git clone --mirror https://github.com/vohoailinh90/Outlook-RSVP-Tool_Event-follow-up-tool.git rsvp-purge.git
cd rsvp-purge.git

git filter-repo --invert-paths --path rsvp_data.db --path RSVP_tool.pdf
```

Kiểm tra: hai lệnh dưới đây phải **không in ra gì**.

```powershell
git log --all --oneline -- rsvp_data.db RSVP_tool.pdf
git rev-list --all --objects | Select-String -Pattern 'rsvp_data\.db|RSVP_tool\.pdf'
```

## 3. Sửa `BASELINE_COMMIT` (bắt buộc, nếu không CI sẽ đỏ)

Viết lại lịch sử làm **mọi SHA commit đều thay đổi**. `scripts/verify_golden_baseline.py`
có ghi cứng `BASELINE_COMMIT = "368c279"` (commit *"Harden the guards against the evasions
that defeated them"*). Sau khi viết lại, SHA đó không còn tồn tại.

Tìm SHA mới:

```powershell
git log --all --oneline --grep "Harden the guards against the evasions"
```

Nếu commit không có trong kết quả (vì PR #1 chưa merge), hãy tra bảng ánh xạ SHA cũ → mới
mà filter-repo ghi lại:

```powershell
Select-String -Path .\filter-repo\commit-map -Pattern '^368c279'
```

Sau khi force-push ở bước 4, sửa SHA đó trong `scripts/verify_golden_baseline.py` trên một
bản clone thường, rồi commit và push như bình thường. Các SHA ghi trong nội dung tài liệu
và commit message chỉ là chữ, sẽ trỏ tới SHA cũ, nhưng không làm hỏng gì.

## 4. Force-push

Nếu nhánh `main` có branch protection chặn force-push, hãy tạm tắt ở Settings → Branches,
rồi bật lại sau khi push xong.

```powershell
git remote add origin https://github.com/vohoailinh90/Outlook-RSVP-Tool_Event-follow-up-tool.git
git push origin --force --all
git push origin --force --tags
```

Không dùng `git push --mirror`. Lệnh này còn cố push `refs/pull/*`, và GitHub từ chối các
ref đó.

## 5. Xoá bản cache trên GitHub

Sau khi force-push, GitHub vẫn giữ các commit cũ: chúng vẫn mở được qua URL chứa SHA và
qua trang diff của PR. Chỉ GitHub Support mới xoá được phần này.

- Liên hệ tại <https://support.github.com/contact>, chọn loại yêu cầu xoá dữ liệu nhạy cảm
  (*removing sensitive data*).
- Cung cấp: tên repo, đường dẫn hai file, SHA các commit cũ đã thêm chúng (`48060db`,
  `9a67145`), và số PR bị ảnh hưởng (#1).
- Tham khảo: GitHub Docs, *Removing sensitive data from a repository*.

## 6. Sau khi xong

- **Mọi bản clone cũ phải xoá và clone lại**, kể cả trên máy bạn và trong các phiên Claude
  Code. Pull hoặc merge từ một bản clone cũ sẽ đưa hai file quay lại lịch sử.
  `scripts/check_no_pii.py` chặn việc file được theo dõi (track) lại, nhưng không nhìn thấy
  lịch sử.
- Chép `rsvp_data.db` từ bản sao lưu vào thư mục của bản clone mới. File này đã nằm trong
  `.gitignore`, nên sẽ không bị commit lại.
- Nếu muốn, chuyển repo về public lại. Hãy làm việc này **sau khi** GitHub Support xác nhận
  đã xoá cache.
- Xoá thư mục `C:\Temp\rsvp-purge.git`.
