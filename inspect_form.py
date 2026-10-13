from pathlib import Path
from playwright.sync_api import sync_playwright

URL = ""
DUMP_FILE = "form_dump.html"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    page.goto(URL, wait_until="domcontentloaded")

    input("Login SITB dan buka menu Data Skrining, lalu ENTER...")

    page.locator("#btnAdd_ta_skrining").wait_for(
        state="visible", timeout=15000
    )
    page.locator("#btnAdd_ta_skrining").click()
    page.wait_for_timeout(2000)

    input(
        "Jika form Tambah sudah terlihat (boleh coba isi NIK & tekan Enter dulu), "
        "lalu tekan ENTER di sini..."
    )

    print("\nURL halaman:", page.url)
    print("Jumlah frame:", len(page.frames))

    for frame_index, frame in enumerate(page.frames):
        print(f"\n{'=' * 20} FRAME {frame_index} {'=' * 20}")
        print("URL frame:", frame.url)

        if "/Skrining/add" in frame.url:
            try:
                html_content = frame.content()
                dump_path = Path(DUMP_FILE)
                dump_path.write_text(html_content, encoding="utf-8")
                print(f"[INFO] Full HTML form berhasil disimpan ke: {dump_path.resolve()}")
            except Exception as e:
                print(f"[PERINGATAN] Gagal menyimpan HTML dump: {e}")

        fields = frame.locator("input, textarea, select, button")

        for i in range(fields.count()):
            field = fields.nth(i)

            try:
                info = field.evaluate("""el => {
                    let kendoRole = el.getAttribute('data-role') || '';
                    if (window.jQuery) {
                        const d = $(el).data();
                        if (d) {
                            const kKeys = Object.keys(d).filter(k => k.toLowerCase().includes('kendo'));
                            if (kKeys.length > 0) kendoRole = kKeys.join(',');
                        }
                    }
                    return {
                        tag: el.tagName,
                        id: el.id,
                        name: el.getAttribute('name'),
                        type: el.getAttribute('type'),
                        value: el.value ?? '',
                        visible: !!(el.offsetWidth || el.offsetHeight),
                        disabled: el.disabled,
                        readonly: el.readOnly,
                        kendo: kendoRole
                    };
                }""")

                if info["visible"] or any(
                    key in str(info).lower()
                    for key in [
                        "dt_tgl", "nik", "nama", "provinsi", "kabupaten",
                        "kecamatan", "kelamin", "lahir", "umur",
                        "berat_badan", "tinggi_badan", "skrining", "save"
                    ]
                ):
                    print(info)

            except Exception as e:
                print("Gagal membaca field:", e)

    input("\nSelesai. Tekan ENTER untuk menutup browser...")
    browser.close()