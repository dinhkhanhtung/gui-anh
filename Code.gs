/**
 * Apps Script GỘP - nhận ảnh khách gửi cho mọi dự án, ghi vào kho chung Client_Projects trên Drive.
 * Cấu trúc: Client_Projects / <thư mục dự án> / Khach_Gui / <nhóm> / <YYYY-MM> / ảnh
 * Dự án mới (mã chưa khai báo): tự tạo Client_Projects/<ma>/Khach_Gui/...
 *
 * Triển khai: Extensions > Apps Script > dán file này > Deploy > Manage deployments >
 * sửa triển khai Web app > Version: New version (Execute as: Me, Access: Anyone).
 */

const ROOT_FOLDER_ID = "1fdgLNCDcNwevoWw7MXYneB8jcdl1NkHb"; // Client_Projects

// Mã dự án (trong link ?p=) -> thư mục dự án đã có sẵn trong Client_Projects
const FOLDER_OF = {
  "ban-dinh": "01_Ban_Dinh_Fanpage",
  "khoan-coc": "02_Khoan_Coc_Nhoi",
  "xe-khanh-quynh": "03_Xe_Ghep_Khanh_Quynh",
  "hoa-nha-tham": "04_Hoa_Nha_Tham_Tran",
  "ha-thao-moc": "05_Ha_Thao_Moc_Ninh_Binh"
};

// Mã hợp lệ = chữ thường, số, gạch ngang (2-40 ký tự). Mã chưa khai báo dùng chính mã làm tên thư mục.
function projFolder_(code) {
  code = String(code || "");
  if (!/^[a-z0-9][a-z0-9\-]{1,39}$/.test(code)) return null;
  return FOLDER_OF[code] || code;
}

function safe_(s) {
  return String(s || "Chung").replace(/[^A-Za-z0-9_\-]/g, "_").slice(0, 60);
}

function sub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Thư mục đích: Client_Projects/<dự án>/Khach_Gui/<nhóm>/<YYYY-MM>
function targetFolder_(proj, category) {
  const month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");
  let folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  folder = sub_(folder, proj);
  folder = sub_(folder, "Khach_Gui");
  folder = sub_(folder, safe_(category));
  return sub_(folder, month);
}

const MAX_VIDEO_BYTES = 500 * 1024 * 1024; // 500 MB / video
const SITE_ORIGIN = "https://dinhkhanhtung.github.io";

// vstart: trả về địa chỉ phiên tải lên (resumable). Trình duyệt PUT từng phần tới địa chỉ này.
function videoStart_(d, proj) {
  const size = Number(d.size || 0);
  const mime = String(d.type || "video/mp4");
  if (mime.indexOf("video/") !== 0) return json_({ status: "error", message: "Chi nhan video" });
  if (!(size > 0) || size > MAX_VIDEO_BYTES) return json_({ status: "error", message: "Video qua lon (toi da 500 MB)" });

  const folder = targetFolder_(proj, d.category);
  const ext = (String(d.name || "").match(/\.(mp4|mov|m4v|webm|3gp|avi|mkv)$/i) || [".mp4"])[0].toLowerCase();
  const name = safe_(String(d.name || "VID_" + Date.now()).replace(/\.[A-Za-z0-9]+$/, "")) + ext;

  const res = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true", {
    method: "post",
    contentType: "application/json; charset=UTF-8",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken(),
      "X-Upload-Content-Type": mime,
      "X-Upload-Content-Length": String(size),
      Origin: SITE_ORIGIN
    },
    payload: JSON.stringify({ name: name, parents: [folder.getId()], mimeType: mime }),
    muteHttpExceptions: true
  });
  const h = res.getHeaders();
  const loc = h["Location"] || h["location"];
  if (res.getResponseCode() !== 200 || !loc) {
    return json_({ status: "error", message: "Khong tao duoc phien tai len: " + res.getResponseCode() + " " + res.getContentText().slice(0, 120) });
  }
  return json_({ status: "success", uploadUrl: loc, name: name });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const proj = projFolder_(d.project);
    if (!proj) return json_({ status: "error", message: "Ma du an khong hop le" });

    // Video: tạo phiên tải lên trực tiếp (resumable) vào Drive, trình duyệt tự gửi từng phần
    if (d.action === "vstart") return videoStart_(d, proj);

    const folder = targetFolder_(proj, d.category);

    const blob = Utilities.newBlob(
      Utilities.base64Decode(d.base64),
      d.type || "image/jpeg",
      safe_((d.name || "IMG_" + Date.now()).replace(/\.jpg$/i, "")) + ".jpg"
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const id = file.getId();
    return json_({
      status: "success",
      id: id,
      url: "https://lh3.googleusercontent.com/d/" + id,
      driveUrl: file.getUrl(),
      project: d.project,
      category: d.category
    });
  } catch (err) {
    return json_({ status: "error", message: err.toString() });
  }
}

// GET ?project=ban-dinh&category=San_Pham -> liệt kê tối đa 60 ảnh khách gửi gần nhất
function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    if (!p.project) return json_({ status: "success", message: "API hoat dong binh thuong." });
    const proj = projFolder_(p.project);
    if (!proj) return json_({ status: "error", message: "Ma du an khong hop le" });

    let folder = sub_(sub_(DriveApp.getFolderById(ROOT_FOLDER_ID), proj), "Khach_Gui");
    if (p.category && p.category !== "all") folder = sub_(folder, safe_(p.category));

    const out = [];
    const walk = (f) => {
      const files = f.getFiles();
      while (files.hasNext() && out.length < 60) {
        const x = files.next();
        if (x.getMimeType().indexOf("image/") === 0) {
          out.push({
            id: x.getId(), name: x.getName(),
            url: "https://lh3.googleusercontent.com/d/" + x.getId(),
            thumb: "https://drive.google.com/thumbnail?id=" + x.getId() + "&sz=w400",
            createdTime: x.getDateCreated().toISOString()
          });
        }
      }
      const subs = f.getFolders();
      while (subs.hasNext() && out.length < 60) walk(subs.next());
    };
    walk(folder);
    return json_({ status: "success", count: out.length, files: out });
  } catch (err) {
    return json_({ status: "error", message: err.toString() });
  }
}

// Chạy 1 lần trong trình soạn thảo để cấp quyền kết nối dịch vụ ngoài (cần cho tải video).
function kiemTraQuyen() {
  const r = UrlFetchApp.fetch("https://www.googleapis.com/discovery/v1/apis?name=drive&preferred=true", { muteHttpExceptions: true });
  Logger.log("External request OK: " + r.getResponseCode());
}
