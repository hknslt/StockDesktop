

/* Timestamp | Date | number -> string formata çevirmek için yardımcı fonksiyon */
function formatlaTarih(v: any): string {
    if (!v) return "-";
    let d: Date | null = null;
    if (typeof v?.toDate === "function") d = v.toDate();
    else if (v instanceof Date) d = v;
    else if (typeof v === "number") d = new Date(v);

    return d ? d.toLocaleDateString("tr-TR") : "-";
}

export function siparisleriExceleAktar(
    liste: any[],
    musteriMap: Record<string, any>,
    etiketMap: Record<string, string>
) {

    const XLSX = (window as any).XLSX;
    const veri: any[][] = [];

    // --- STİL TANIMLAMALARI ---
    const musteriBaslikStili = {
        font: { bold: true, color: { rgb: "FFFFFF" } }, // Beyaz yazı
        fill: { fgColor: { rgb: "2F5597" } }, // Koyu Mavi arka plan
        alignment: { vertical: "center", horizontal: "center" },
        border: {
            top: { style: "thin", color: { auto: 1 } },
            bottom: { style: "thin", color: { auto: 1 } },
            left: { style: "thin", color: { auto: 1 } },
            right: { style: "thin", color: { auto: 1 } }
        }
    };

    const urunBaslikStili = {
        font: { bold: true, color: { rgb: "000000" } }, // Siyah yazı
        fill: { fgColor: { rgb: "D9E1F2" } }, // Açık Mavi arka plan
        alignment: { vertical: "center", horizontal: "center" },
        border: {
            top: { style: "thin", color: { auto: 1 } },
            bottom: { style: "thin", color: { auto: 1 } },
            left: { style: "thin", color: { auto: 1 } },
            right: { style: "thin", color: { auto: 1 } }
        }
    };

    const hucreKenarlikStili = {
        border: {
            bottom: { style: "thin", color: { rgb: "DDDDDD" } }
        }
    };

    liste.forEach((r) => {
        const idKey = String(r.musteri?.id || "");
        const m = musteriMap[idKey] || r.musteri || {};
        const musteriAd = m.firmaAdi || m.yetkili || "-";
        const durum = etiketMap[r.durum] || r.durum;

        // 1. Müşteri/Sipariş Satırı Başlıkları (Stilli Hücre Objeleri)
        veri.push([
            { v: "MÜŞTERİ", t: "s", s: musteriBaslikStili },
            { v: "TEL", t: "s", s: musteriBaslikStili },
            { v: "ADRES", t: "s", s: musteriBaslikStili },
            { v: "TARİH", t: "s", s: musteriBaslikStili },
            { v: "İŞLEM TARİHİ", t: "s", s: musteriBaslikStili },
            { v: "SİPARİŞ TUTARI (BRÜT)", t: "s", s: musteriBaslikStili },
            { v: "DURUMU", t: "s", s: musteriBaslikStili },
        ]);

        // 2. Müşteri/Sipariş Verileri
        veri.push([
            { v: musteriAd, t: "s", s: hucreKenarlikStili },
            { v: m.telefon || "-", t: "s", s: hucreKenarlikStili },
            { v: m.adres || "-", t: "s", s: hucreKenarlikStili },
            { v: formatlaTarih(r.tarih), t: "s", s: hucreKenarlikStili },
            { v: formatlaTarih(r.islemeTarihi), t: "s", s: hucreKenarlikStili },
            { v: Number(r.brutTutar || 0), t: "n", s: hucreKenarlikStili },
            { v: durum, t: "s", s: hucreKenarlikStili },
        ]);

        // 3. Ürün Satırı Başlıkları
        veri.push([
            "", // İlk sütun boş (İçerik hiyerarşisi için)
            { v: "ÜRÜN ADI VE RENGİ", t: "s", s: urunBaslikStili },
            { v: "ADET", t: "s", s: urunBaslikStili },
            { v: "BİRİM FİYAT (NET)", t: "s", s: urunBaslikStili },
            { v: "ÜRÜN TOPLAM FİYAT", t: "s", s: urunBaslikStili },
        ]);

        // 4. Ürün Verileri
        const urunler = r.urunler || [];
        urunler.forEach((u: any) => {
            const urunAdRenk = u.renk ? `${u.urunAdi} - ${u.renk}` : u.urunAdi;
            const adet = Number(u.adet || 0);
            const birimFiyat = Number(u.birimFiyat || 0);
            const toplam = adet * birimFiyat;

            veri.push([
                "",
                { v: urunAdRenk, t: "s" },
                { v: adet, t: "n" },
                { v: birimFiyat, t: "n" },
                { v: toplam, t: "n" }
            ]);
        });

        // Her siparişin arasına okunaklı olması için iki boş satır
        veri.push([]);
        veri.push([]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(veri);

    // Sütun Genişlikleri
    worksheet["!cols"] = [
        { wch: 30 }, // A Sütunu: Müşteri
        { wch: 35 }, // B Sütunu: Tel / Ürün Adı ve Rengi
        { wch: 45 }, // C Sütunu: Adres / Adet
        { wch: 15 }, // D Sütunu: Tarih / Birim Fiyat
        { wch: 18 }, // E Sütunu: İşlem Tarihi / Toplam
        { wch: 22 }, // F Sütunu: Sipariş Tutarı
        { wch: 15 }, // G Sütunu: Durumu
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Siparişler");
    XLSX.writeFile(workbook, "Siparis_Listesi.xlsx");
}