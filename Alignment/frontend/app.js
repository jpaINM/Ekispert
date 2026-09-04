(() => {
  "use strict";

  const committeeListEl = document.getElementById("committee-list");
  const committeeListEmpty = document.getElementById("committee-list-empty");
  const memberTabsEl = document.getElementById("member-tabs");
  const memberEmpty = document.getElementById("member-empty");
  const memberDetailEl = document.getElementById("member-detail");
  const selectedCommitteeName = document.getElementById("selected-committee-name");

  // ---- カスタム日付ピッカー ----
  const datePickerEl = document.getElementById("date-picker");
  const dateTriggerEl = document.getElementById("date-trigger");
  const dateTriggerTextEl = document.getElementById("date-trigger-text");
  const dateCalendarEl = document.getElementById("date-calendar");
  const calTitleEl = document.getElementById("cal-title");
  const calGridEl = document.getElementById("cal-grid");
  const calPrevEl = document.getElementById("cal-prev");
  const calNextEl = document.getElementById("cal-next");

  let selectedDate = new Date();      // 選択日（デフォルト=今日）
  let viewYear = selectedDate.getFullYear();
  let viewMonth = selectedDate.getMonth();  // 0-11
  let calendarOpen = false;

  const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  // YYYY-MM-DD（プロキシ互換）
  function formatDateISO(date) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  // YYYY/MM/DD（トリガー表示用）
  function formatDateDisplay(date) {
    return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(date.getDate())}`;
  }

  function isSameDay(a, b) {
    return a.getFullYear() === b.getFullYear()
      && a.getMonth() === b.getMonth()
      && a.getDate() === b.getDate();
  }

  function updateTriggerText() {
    if (dateTriggerTextEl) {
      dateTriggerTextEl.textContent = formatDateDisplay(selectedDate);
    }
  }

  function updateCalendarTitle() {
    if (calTitleEl) {
      calTitleEl.textContent = `${viewYear}年${viewMonth + 1}月`;
    }
  }

  function renderCalendar() {
    if (!calGridEl) return;
    calGridEl.innerHTML = "";

    const firstDay = new Date(viewYear, viewMonth, 1);
    const startWeekday = firstDay.getDay();          // 0=日
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
    const daysInPrevMonth = new Date(viewYear, viewMonth, 0).getDate();

    const today = new Date();

    // 前月の余白
    for (let i = startWeekday - 1; i >= 0; i--) {
      const day = daysInPrevMonth - i;
      calGridEl.appendChild(createCell(day, true, viewYear, viewMonth - 1, today));
    }

    // 当月
    for (let day = 1; day <= daysInMonth; day++) {
      calGridEl.appendChild(createCell(day, false, viewYear, viewMonth, today));
    }

    // 翌月の余白（42マス=6行になるまで埋める）
    const filled = startWeekday + daysInMonth;
    const remain = (7 - (filled % 7)) % 7;
    for (let day = 1; day <= remain; day++) {
      calGridEl.appendChild(createCell(day, true, viewYear, viewMonth + 1, today));
    }
  }

  function createCell(day, isOutside, year, month, today) {
    // month が -1 や 12 の場合は正規化
    const cellDate = new Date(year, month, day);
    const realYear = cellDate.getFullYear();
    const realMonth = cellDate.getMonth();

    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = "calendar-cell";
    cell.textContent = String(day);

    if (isOutside) cell.classList.add("is-outside");
    if (isSameDay(cellDate, today)) cell.classList.add("is-today");
    if (isSameDay(cellDate, selectedDate)) cell.classList.add("is-selected");

    cell.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedDate = new Date(realYear, realMonth, day);
      viewYear = realYear;
      viewMonth = realMonth;
      updateTriggerText();
      renderCalendar();
      closeCalendar();
    });

    return cell;
  }

  function openCalendar() {
    if (!dateCalendarEl) return;
    calendarOpen = true;
    dateCalendarEl.hidden = false;
    dateTriggerEl.setAttribute("aria-expanded", "true");
    // 選択月を表示対象に同期
    viewYear = selectedDate.getFullYear();
    viewMonth = selectedDate.getMonth();
    updateCalendarTitle();
    renderCalendar();
  }

  function closeCalendar() {
    if (!dateCalendarEl) return;
    calendarOpen = false;
    dateCalendarEl.hidden = true;
    dateTriggerEl.setAttribute("aria-expanded", "false");
  }

  function toggleCalendar() {
    if (calendarOpen) closeCalendar();
    else openCalendar();
  }

  function shiftMonth(delta) {
    viewMonth += delta;
    if (viewMonth < 0) {
      viewMonth = 11;
      viewYear -= 1;
    } else if (viewMonth > 11) {
      viewMonth = 0;
      viewYear += 1;
    }
    updateCalendarTitle();
    renderCalendar();
  }

  function initDatePicker() {
    updateTriggerText();
    if (dateTriggerEl) {
      dateTriggerEl.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleCalendar();
      });
    }
    if (calPrevEl) calPrevEl.addEventListener("click", (e) => {
      e.stopPropagation();
      shiftMonth(-1);
    });
    if (calNextEl) calNextEl.addEventListener("click", (e) => {
      e.stopPropagation();
      shiftMonth(1);
    });
    // 外部クリックで閉じる
    document.addEventListener("click", (e) => {
      if (!calendarOpen) return;
      if (datePickerEl && datePickerEl.contains(e.target)) return;
      closeCalendar();
    });
    // Esc で閉じる
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && calendarOpen) closeCalendar();
    });
  }

  // 出発日を取得（後続フェーズのプロキシ実装で使用）
  function getSearchDate() {
    return formatDateISO(selectedDate);
  }

  let activeCommittee = null;
  let activeMemberIndex = 0;
  let activeCommitteeButton = null;

  // データ取得(DB参照APIから取得)
  async function fetchData() {
    const res = await fetch("/api/committees");
    if (!res.ok) throw new Error(`APIエラー: ${res.status}`);
    const json = await res.json();
    if (!json.ok) throw new Error(json.error?.message || "API取得失敗");
    return json;
  }

  function renderCommitteeList(data) {
    committeeListEl.innerHTML = "";

    if (!data.fiscal_years || !data.fiscal_years.length) {
      committeeListEl.appendChild(committeeListEmpty);
      committeeListEmpty.textContent = "委員会がありません";
      return;
    }

    data.fiscal_years.forEach((fiscalYear) => {
      const group = document.createElement("div");
      group.className = "fiscal-year-group";

      const label = document.createElement("p");
      label.className = "fiscal-year-label";
      label.textContent = fiscalYear.label;
      group.appendChild(label);

      fiscalYear.committees.forEach((committee) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "committee-item";
        button.title = committee.name;

        const dot = document.createElement("span");
        dot.className = "committee-dot";

        const name = document.createElement("span");
        name.className = "committee-name";
        name.textContent = committee.name;

        button.append(dot, name);
        button.addEventListener("click", () =>
          selectCommittee(committee, button, fiscalYear.label)
        );
        group.appendChild(button);
      });

      committeeListEl.appendChild(group);
    });
  }

  function clearActiveCommittee() {
    if (activeCommitteeButton) {
      activeCommitteeButton.classList.remove("is-active");
      activeCommitteeButton = null;
    }
  }

  function selectCommittee(committee, button, fiscalYearLabel) {
    clearActiveCommittee();
    button.classList.add("is-active");
    activeCommitteeButton = button;
    activeCommittee = committee;
    activeMemberIndex = 0;

    selectedCommitteeName.hidden = false;
    selectedCommitteeName.textContent = `${fiscalYearLabel} ${committee.name}`;

    renderMemberTabs(committee);
    if (committee.members && committee.members.length) {
      renderMemberDetail(committee.members[0]);
    } else {
      renderMemberEmpty("この委員会にはメンバーがいません");
    }
  }

  function renderMemberTabs(committee) {
    memberTabsEl.innerHTML = "";

    if (!committee.members || !committee.members.length) {
      memberTabsEl.appendChild(memberEmpty);
      memberEmpty.textContent = "メンバーがいません";
      return;
    }

    committee.members.forEach((member, index) => {
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "member-tab";
      tab.setAttribute("role", "tab");
      tab.textContent = member.name;
      if (index === activeMemberIndex) tab.classList.add("is-active");
      tab.addEventListener("click", () => {
        activeMemberIndex = index;
        document.querySelectorAll(".member-tab").forEach((t) =>
          t.classList.remove("is-active")
        );
        tab.classList.add("is-active");
        renderMemberDetail(member);
      });
      memberTabsEl.appendChild(tab);
    });
  }

  function renderMemberDetail(member) {
    const vias = member.via_stations || [];

    const viaSteps = vias
      .map(
        (station, i) => `
        <div class="route-step is-via">
          <span class="route-marker">経</span>
          <span class="route-station">${escapeHtml(station)}</span>
        </div>`
      )
      .join("");

    memberDetailEl.innerHTML = `
      <div class="detail-grid">
        <span class="detail-label">役職</span>
        <p class="detail-value">${escapeHtml(member.position || "")}</p>
        <span class="detail-label">氏名</span>
        <p class="detail-value">${escapeHtml(member.name || "")}</p>
      </div>
      <div class="route-section">
        <p class="route-title">交通経路</p>
        <div class="route-flow">
          <div class="route-step">
            <span class="route-marker">発</span>
            <span class="route-station">${escapeHtml(member.departure_station || "")}</span>
          </div>
          ${viaSteps}
          <div class="route-step">
            <span class="route-marker">着</span>
            <span class="route-station">${escapeHtml(member.arrival_station || "")}</span>
          </div>
        </div>
      </div>`;
  }

  function renderMemberEmpty(message) {
    memberDetailEl.innerHTML = `<p class="detail-empty">${escapeHtml(message)}</p>`;
  }

  function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = String(text);
    return div.innerHTML;
  }

  // 駅すぱあと連携ボタン(モック: クリックでアニメーション開始、3秒後に完了で停止)
  const ekispertButton = document.getElementById("ekispert-sync-button");
  let ekispertRunning = false;
  let ekispertTimer = null;

  function setEkispertRunning(running) {
    ekispertRunning = running;
    ekispertButton.dataset.running = running ? "true" : "false";
    ekispertButton.disabled = running;
    ekispertButton.textContent = running ? "連携中..." : "駅すぱあと連携";
  }

  if (ekispertButton) ekispertButton.addEventListener("click", () => {
    if (ekispertRunning) return;
    setEkispertRunning(true);
    console.log("[Ekispert] 連携開始(モック) 出発日:", getSearchDate());

    // モック完了: 3秒後に停止(実際のAPI連携は後続フェーズで差し替え)
    ekispertTimer = setTimeout(() => {
      setEkispertRunning(false);
      console.log("[Ekispert] 連携完了(モック)");
    }, 3000);
  });

  async function initialize() {
    initDatePicker();
    try {
      const data = await fetchData();
      renderCommitteeList(data);
    } catch (error) {
      committeeListEl.appendChild(committeeListEmpty);
      committeeListEmpty.textContent = "データの読み込みに失敗しました";
    }
  }

  initialize();
})();