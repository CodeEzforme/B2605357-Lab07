const LAB07 = {
  closeDate: { year: 2026, month: 12, day: 1 },
  timezone: 'Asia/Bangkok',
  tmaLimit: 80,
};

function buildLab07() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAB07_CREATED') === 'true') {
    throw new Error('Bộ Lab 07 đã được tạo. Mở Project Settings > Script properties để xem ID.');
  }

  const studentForm = createStudentForm_();
  const tmaForm = createTmaForm_();
  const calendar = createOctoberPlan_();

  props.setProperties({
    LAB07_CREATED: 'true',
    STUDENT_FORM_ID: studentForm.getId(),
    STUDENT_FORM_URL: studentForm.getPublishedUrl(),
    STUDENT_FORM_EDIT_URL: studentForm.getEditUrl(),
    TMA_FORM_ID: tmaForm.getId(),
    TMA_FORM_URL: tmaForm.getPublishedUrl(),
    TMA_FORM_EDIT_URL: tmaForm.getEditUrl(),
    CALENDAR_ID: calendar.getId(),
  });

  ScriptApp.newTrigger('closeStudentForm')
    .timeBased()
    .atDate(LAB07.closeDate.year, LAB07.closeDate.month, LAB07.closeDate.day)
    .inTimezone(LAB07.timezone)
    .create();

  ScriptApp.newTrigger('handleTmaSubmit')
    .forForm(tmaForm)
    .onFormSubmit()
    .create();

  console.log(JSON.stringify(getLab07Links(), null, 2));
}

function createStudentForm_() {
  const form = FormApp.create('Thông tin sinh viên khóa 52');
  form
    .setDescription('Vui lòng nhập đầy đủ và chính xác thông tin sinh viên. Biểu mẫu ngưng nhận phản hồi từ ngày 01/12/2026.')
    .setCollectEmail(false)
    .setProgressBar(true)
    .setConfirmationMessage('Thông tin của bạn đã được ghi nhận.')
    .setCustomClosedFormMessage('Biểu mẫu đã ngưng nhận phản hồi từ ngày 01/12/2026.');
  publishIfSupported_(form);

  addText_(form, 'Mã sinh viên');
  addText_(form, 'Họ và tên');
  addGender_(form);
  form.addDateItem().setTitle('Ngày sinh').setRequired(true);
  addText_(form, 'Nơi sinh');
  addText_(form, 'Ngành');
  addText_(form, 'Mã lớp');
  addText_(form, 'Họ tên giáo viên cố vấn');
  addText_(form, 'Email');
  addText_(form, 'Điện thoại');
  addText_(form, 'Họ tên cha');
  addText_(form, 'Số điện thoại cha');
  addText_(form, 'Nghề nghiệp của cha');
  addText_(form, 'Họ tên mẹ');
  addText_(form, 'Số điện thoại mẹ');
  addText_(form, 'Nghề nghiệp của mẹ');
  form.addParagraphTextItem().setTitle('Địa chỉ liên hệ gia đình').setRequired(true);

  const deadline = new Date(2026, 11, 1, 0, 0, 0);
  if (new Date() >= deadline) form.setAcceptingResponses(false);
  return form;
}

function createTmaForm_() {
  const form = FormApp.create('Đăng ký tham quan Công ty TMA tại TP. Hồ Chí Minh');
  form
    .setDescription('Dành cho sinh viên khóa 51. Biểu mẫu tự ngưng nhận phản hồi khi đủ 80 sinh viên đăng ký.')
    .setCollectEmail(false)
    .setProgressBar(true)
    .setConfirmationMessage('Đăng ký thành công. Ban tổ chức sẽ liên hệ qua email hoặc điện thoại.')
    .setCustomClosedFormMessage('Đợt đăng ký đã đủ 80 sinh viên và hiện đã đóng.');
  publishIfSupported_(form);

  addText_(form, 'Mã sinh viên');
  addText_(form, 'Họ và tên');
  addGender_(form);
  form.addDateItem().setTitle('Ngày sinh').setRequired(true);
  addText_(form, 'Nơi sinh');
  addText_(form, 'Ngành');
  addText_(form, 'Mã lớp');
  addText_(form, 'Họ tên giáo viên cố vấn');
  addText_(form, 'Email');
  addText_(form, 'Điện thoại');
  return form;
}

function publishIfSupported_(form) {
  if (form.supportsAdvancedResponderPermissions()) form.setPublished(true);
  form.setAcceptingResponses(true);
}

function addText_(form, title) {
  form.addTextItem().setTitle(title).setRequired(true);
}

function addGender_(form) {
  form.addMultipleChoiceItem()
    .setTitle('Phái')
    .setChoiceValues(['Nam', 'Nữ', 'Khác'])
    .setRequired(true);
}

function closeStudentForm() {
  const id = PropertiesService.getScriptProperties().getProperty('STUDENT_FORM_ID');
  if (!id) throw new Error('Không tìm thấy STUDENT_FORM_ID. Hãy chạy buildLab07 trước.');
  FormApp.openById(id)
    .setCustomClosedFormMessage('Biểu mẫu đã ngưng nhận phản hồi từ ngày 01/12/2026.')
    .setAcceptingResponses(false);
}

function handleTmaSubmit() {
  const id = PropertiesService.getScriptProperties().getProperty('TMA_FORM_ID');
  if (!id) throw new Error('Không tìm thấy TMA_FORM_ID. Hãy chạy buildLab07 trước.');
  const form = FormApp.openById(id);
  if (form.getResponses().length >= LAB07.tmaLimit) {
    form
      .setCustomClosedFormMessage('Đợt đăng ký đã đủ 80 sinh viên và hiện đã đóng.')
      .setAcceptingResponses(false);
  }
}

function createOctoberPlan_() {
  const calendar = CalendarApp.createCalendar('Kế hoạch cá nhân tháng 10 năm 2026', {
    summary: 'Kế hoạch học tập, bài tập, rèn luyện và tổng kết theo tuần.',
    timeZone: LAB07.timezone,
  });
  const events = [
    [1, 19, 'Lập kế hoạch tháng', 'Xác định mục tiêu học tập và ba ưu tiên chính.'],
    [2, 19, 'Ôn chuyên ngành', 'Ôn lại nội dung đã học trong tuần.'],
    [4, 20, 'Rà soát tuần 1', 'Kiểm tra việc đã hoàn thành và cập nhật Google Tasks.'],
    [5, 19, 'Học lập trình', 'Luyện tập bài tập lập trình trong 90 phút.'],
    [6, 19, 'Bài tập nhóm', 'Cập nhật tiến độ và chia việc cho nhóm.'],
    [8, 19, 'Đọc tài liệu', 'Đọc tài liệu chuyên ngành và ghi chú bằng Google Keep.'],
    [11, 20, 'Rà soát tuần 2', 'Tổng kết kết quả tuần và điều chỉnh ưu tiên.'],
    [12, 19, 'Học lập trình', 'Hoàn thành bài thực hành đang làm.'],
    [14, 19, 'Họp nhóm', 'Kiểm tra tiến độ và xử lý vướng mắc.'],
    [16, 19, 'Hoàn thiện bài tập', 'Chuẩn bị bản nộp trước hạn.'],
    [18, 20, 'Rà soát tuần 3', 'Cập nhật các việc còn tồn.'],
    [19, 19, 'Ôn giữa kỳ', 'Hệ thống hóa kiến thức theo chủ đề.'],
    [21, 19, 'Bài tập nhóm', 'Tích hợp phần việc của các thành viên.'],
    [23, 19, 'Nộp bài đúng hạn', 'Kiểm tra tệp và nộp bài trước giờ chót.'],
    [25, 20, 'Rà soát tuần 4', 'Đối chiếu mục tiêu tháng.'],
    [26, 19, 'Tổng hợp ghi chú', 'Sắp xếp ghi chú học tập theo môn.'],
    [28, 19, 'Hoàn thiện Lab 07', 'Kiểm tra biểu mẫu, website, kế hoạch và báo cáo.'],
    [30, 19, 'Đánh giá tháng', 'Ghi lại kết quả và điểm cần cải thiện.'],
    [31, 9, 'Lập kế hoạch tháng mới', 'Chuẩn bị mục tiêu cho tháng 11/2026.'],
  ];
  events.forEach(([day, hour, title, description]) => {
    const start = new Date(2026, 9, day, hour, 0, 0);
    const end = new Date(2026, 9, day, hour + 1, 0, 0);
    calendar.createEvent(title, start, end, { description }).addPopupReminder(30);
  });
  return calendar;
}

function getLab07Links() {
  const props = PropertiesService.getScriptProperties();
  return {
    studentForm: props.getProperty('STUDENT_FORM_URL'),
    studentFormEdit: props.getProperty('STUDENT_FORM_EDIT_URL'),
    tmaForm: props.getProperty('TMA_FORM_URL'),
    tmaFormEdit: props.getProperty('TMA_FORM_EDIT_URL'),
    calendarId: props.getProperty('CALENDAR_ID'),
  };
}
