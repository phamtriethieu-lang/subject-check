/* Subject Check - Outlook add-in (Smart Alerts / OnMessageSend)
   Quy dinh: tieu de = yyyyMMdd - Tieu de ; bo "RE:" / "FW:" / "Fwd:" khi reply / forward. */

var PREFIX_RE = /^\s*(re|fw|fwd|tl|ct|tr|aw|wg|rv|sv|vs|r|f)\s*(\[\d+\])?\s*[:：]\s*/i;
var OLDDATE_RE = /^\s*\d{8}\s*[-–]\s*/;

function cleanSubject(s) {
  var prev;
  s = (s || "").trim();
  do {
    prev = s;
    s = s.replace(PREFIX_RE, "").replace(OLDDATE_RE, "").trim();
  } while (s !== prev);
  return s;
}

function pad2(n) { return (n < 10 ? "0" : "") + n; }

function todayStr() {
  var d = new Date();
  return "" + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate());
}

function expectedSubject(old) {
  var core = cleanSubject(old);
  return core ? todayStr() + " - " + core : "";
}

/* Chay khi bam Send */
function onMessageSendHandler(event) {
  var done = false;
  function finish(opts) {
    if (done) return;
    done = true;
    event.completed(opts);
  }
  /* An toan: neu qua 8 giay chua xong thi cho gui, khong treo mail */
  var timer = setTimeout(function () { finish({ allowEvent: true }); }, 8000);
  try {
    checkSubject(Office.context.mailbox.item, function (opts) { clearTimeout(timer); finish(opts); });
  } catch (e) {
    clearTimeout(timer);
    finish({ allowEvent: true });
  }
}

function checkSubject(item, complete) {
  var event = { completed: complete };
  item.subject.getAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }
    var old = (r.value || "").trim();
    var nw = expectedSubject(old);

    if (!nw) {
      event.completed({
        allowEvent: false,
        errorMessage: "Email chưa có tiêu đề. Quy định: yyyyMMdd - Tiêu đề (ví dụ: 20261110 - JIG UPDATE)."
      });
      return;
    }
    if (nw === old) {
      event.completed({ allowEvent: true });
      return;
    }

    item.subject.setAsync(nw, function (s) {
      if (s.status === Office.AsyncResultStatus.Succeeded) {
        event.completed({
          allowEvent: false,
          errorMessage: "Tiêu đề chưa đúng quy định, đã sửa thành: " + nw +
            ". Kiểm tra lại rồi bấm Send lần nữa để gửi."
        });
      } else {
        event.completed({
          allowEvent: false,
          errorMessage: "Tiêu đề chưa đúng quy định. Hãy sửa thành: " + nw
        });
      }
    });
  });
}

/* Nut "Fix subject" tren thanh cong cu khi soan mail */
function fixSubject(event) {
  var item = Office.context.mailbox.item;
  item.subject.getAsync(function (r) {
    var nw = expectedSubject(r.value);
    if (!nw) { event.completed(); return; }
    item.subject.setAsync(nw, function () { event.completed(); });
  });
}

/* Outlook moi / Outlook web can Office.onReady de khoi dong add-in */
try { if (typeof Office.onReady === "function") { Office.onReady(function () {}); } } catch (e) {}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("fixSubject", fixSubject);
