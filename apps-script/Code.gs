const NOTIFY_EMAIL = "shunyoutou@gmail.com";
const BOOKINGS_KEY = "bookings";

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const booking = {
      date: sanitize(payload.date),
      time: sanitize(payload.time),
      name: sanitize(payload.name),
      contact: sanitize(payload.contact),
      message: sanitize(payload.message || "未填写"),
      page: sanitize(payload.page || ""),
    };

    if (!booking.date || !booking.time || !booking.name || !booking.contact) {
      return jsonResponse({ ok: false, message: "missing required fields" });
    }

    const key = makeBookingKey(booking.date, booking.time);
    const bookings = getBookings();

    if (bookings[key]) {
      return jsonResponse({ ok: false, message: "already booked" });
    }

    bookings[key] = {
      date: booking.date,
      time: booking.time,
      createdAt: new Date().toISOString(),
    };
    saveBookings(bookings);

    const subject = `MooM预约申请：${booking.date} ${booking.time}`;
    const body = [
      "收到新的预约申请：",
      "",
      `日期：${booking.date}`,
      `时间：${booking.time}`,
      `预约人：${booking.name}`,
      `联系方式：${booking.contact}`,
      `预约内容：${booking.message}`,
      "",
      booking.page ? `来源页面：${booking.page}` : "",
      "",
      "请及时确认预约。",
    ].filter(Boolean).join("\n");

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject,
      body,
      name: "MooM预约通知",
    });

    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ ok: false, message: String(error) });
  }
}

function doGet(e) {
  const bookings = Object.values(getBookings());
  const callback = e && e.parameter && e.parameter.callback;
  const data = { ok: true, bookings };

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(data)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(data);
}

function sanitize(value) {
  return String(value || "").replace(/[<>]/g, "").trim().slice(0, 1000);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function makeBookingKey(date, time) {
  return `${date}|${time}`;
}

function getBookings() {
  const raw = PropertiesService.getScriptProperties().getProperty(BOOKINGS_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveBookings(bookings) {
  PropertiesService.getScriptProperties().setProperty(BOOKINGS_KEY, JSON.stringify(bookings));
}
