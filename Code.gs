/**
 * Apps Script GỘP - nhận ảnh cho mọi dự án, tự chia thư mục Drive.
 * Cấu trúc: KHO_ANH_DU_AN / <project> / <category> / <YYYY-MM> / ảnh
 *
 * Triển khai: Extensions > Apps Script > dán file này > điền ROOT_FOLDER_ID
 * > Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 * Sau đó dán link /exec vào "endpoint" trong du_an.json.
 */

const ROOT_FOLDER_ID = "1xms7jONmNTjkzjQDsBIhYJl7lJrROueB";

// Khách/dự án mới: KHÔNG cần khai báo ở đây. Mã hợp lệ = chữ thường, số, gạch ngang (2-40 ký tự).
// Thư mục Drive tự tạo lần gửi ảnh đầu tiên, tên thư mục = mã.
function projFolder_(code) {
  code = String(code || "");
  return /^[a-z0-9][a-z0-9\-]{1,39}$/.test(code) ? code : null;
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

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const proj = projFolder_(d.project);
    if (!proj) return json_({ status: "error", message: "Ma du an khong hop le" });

    const tz = Session.getScriptTimeZone();
    const month = Utilities.formatDate(new Date(), tz, "yyyy-MM");
    let folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    folder = sub_(folder, proj);
    folder = sub_(folder, safe_(d.category));
    folder = sub_(folder, month);

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

// GET ?project=ban-dinh&category=San_Pham -> liệt kê tối đa 60 ảnh gần nhất trong nhóm
function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    if (!p.project) return json_({ status: "success", message: "API hoat dong binh thuong." });
    const proj = projFolder_(p.project);
    if (!proj) return json_({ status: "error", message: "Ma du an khong hop le" });

    let folder = sub_(DriveApp.getFolderById(ROOT_FOLDER_ID), proj);
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
