const SHEET_HASIL = 'HASILUJIAN';
const SHEET_ESAI = 'DETAIL ESAY';
const SHEET_KUNCI = 'KUNCI PG';
const NILAI_MINIMUM = 80;
const MAKS_ESAI = 50;
const GEMINI_MODEL = 'gemini-2.5-flash';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Ujian')
    .setTitle('ViskoLab | Ujian')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getExam(zone) {
  zone = Number(zone || 1);
  const banks = getQuestionBanks();
  return banks[zone] || banks[1];
}

function submitExam(payload) {
  const zone = Number(payload.zone || 1);
  const nama = String(payload.nama || '').trim();
  const kelas = String(payload.kelas || '').trim();
  if (!nama || !kelas) throw new Error('Nama dan kelas wajib diisi.');

  const bank = getExam(zone);
  const answers = payload.answers || {};
  let pgScore = 0;
  const wrongTopics = [];
  const mastered = [];
  const pgRows = [];

  bank.objective.forEach((q, i) => {
    const key = 'q' + i;
    const user = answers[key];
    let correct = false;
    if (q.type === 'multi') {
      const a = Array.isArray(user) ? user.map(String).sort() : [];
      const b = q.answer.map(String).sort();
      correct = JSON.stringify(a) === JSON.stringify(b);
    } else {
      correct = String(user ?? '') === String(q.answer);
    }
    if (correct) {
      pgScore += q.points;
      mastered.push(q.topic);
    } else {
      wrongTopics.push(q.topic);
    }
    pgRows.push({question:q.question, answer:user, correct, topic:q.topic, points:q.points});
  });

  const essayInput = bank.essay.map((q, i) => ({
    nomor: i + 1,
    topic: q.topic,
    question: q.question,
    answer: String(answers['e' + i] || '').trim(),
    max: q.points,
    rubric: q.rubric
  }));

  const ai = gradeEssaysWithAI(zone, essayInput);
  const essayScore = ai.totalScore;
  ai.items.forEach(item => {
    if (item.score < item.max) wrongTopics.push(item.topic);
    else mastered.push(item.topic);
  });

  const total = Math.round(pgScore + essayScore);
  const passed = total >= NILAI_MINIMUM;
  const uniqueWeak = [...new Set(wrongTopics)];
  const uniqueMastered = [...new Set(mastered)];
  const reviewTopics = uniqueWeak.map(t => {
    const q = bank.objective.find(x => x.topic === t) || bank.essay.find(x => x.topic === t);
    return t + ': ' + (q && q.review ? q.review : 'Pelajari kembali konsep terkait di Pahami dan ulangi pengamatan pada simulasi.');
  });

  saveResult(zone, nama, kelas, pgScore, essayInput, ai, total, passed, pgRows);

  return {
    score: total,
    passed,
    mastered: uniqueMastered,
    weaknesses: uniqueWeak,
    reviewTopics,
    feedback: passed
      ? 'Selamat! Nilaimu sudah mencapai batas minimal 80. Tetap periksa feedback untuk melihat konsep yang sudah kuat.'
      : 'Nilaimu belum mencapai 80. Perbaiki konsep yang tercantum pada feedback, pelajari kembali materi, lalu kerjakan ujian lagi.',
    pgScore,
    essayScore,
    essayFeedback: ai.items,
    zone
  };
}

function gradeEssaysWithAI(zone, essays) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY belum disetel di Script Properties Apps Script.');

  const prompt = `Kamu adalah penilai esai Fisika SMA untuk media ViskoLab. Materi hanya tentang viskositas fluida, suhu, aliran laminar, hukum Newton tentang viskositas, hukum Stokes, kecepatan terminal, Poiseuille, data/grafik, eksperimen, dan penerapan.\n\nNilai setiap jawaban berdasarkan rubrik. Jangan menilai gaya bahasa; nilai ketepatan konsep dan alasan fisikanya. Jika jawaban kosong atau tidak menjawab, beri 0. Berikan feedback singkat dan spesifik dalam bahasa Indonesia.\n\nUjian zona: ${zone}\nData esai:\n${JSON.stringify(essays)}\n\nKembalikan JSON SAJA dengan format persis:\n{"items":[{"nomor":1,"topic":"...","score":0,"max":10,"feedback":"..."}],"totalScore":0}`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({contents:[{parts:[{text:prompt}]}], generationConfig:{temperature:0.1, responseMimeType:'application/json'}})
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('AI gagal menilai esai. Kode API: ' + code);
  const data = JSON.parse(res.getContentText());
  const text = data.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  const clean = text.replace(/^```json\s*/i,'').replace(/```$/,'').trim();
  const parsed = JSON.parse(clean);
  const items = (parsed.items || []).map((x, i) => ({
    nomor: i + 1,
    topic: x.topic || essays[i]?.topic || 'Esai ' + (i + 1),
    score: Math.max(0, Math.min(Number(x.score || 0), Number(essays[i]?.max || 10))),
    max: Number(essays[i]?.max || 10),
    feedback: String(x.feedback || 'Belum ada feedback.')
  }));
  const totalScore = Math.round(items.reduce((s,x) => s + x.score, 0));
  return {items, totalScore};
}

function saveResult(zone, nama, kelas, pgScore, essays, ai, total, passed, pgRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = ss.getSheetByName(SHEET_HASIL) || ss.insertSheet(SHEET_HASIL);
  const detail = ss.getSheetByName(SHEET_ESAI) || ss.insertSheet(SHEET_ESAI);
  const kunci = ss.getSheetByName(SHEET_KUNCI) || ss.insertSheet(SHEET_KUNCI);
  ensureHeaders(hasil, ['Waktu','Zona','Nama','Kelas','Nilai PG','Nilai Esai','Nilai Akhir','Status']);
  ensureHeaders(detail, ['Waktu','Zona','Nama','Kelas','Nomor Soal','Pertanyaan','Jawaban','Skor','Maksimal','Feedback AI']);
  ensureHeaders(kunci, ['Zona','Nomor','Topik','Pertanyaan','Kunci']);
  const now = new Date();
  hasil.appendRow([now, zone, nama, kelas, pgScore, ai.totalScore, total, passed ? 'LULUS' : 'BELUM LULUS']);
  essays.forEach((e, i) => {
    const a = ai.items[i];
    detail.appendRow([now, zone, nama, kelas, e.nomor, e.question, e.answer, a?.score || 0, e.max, a?.feedback || '']);
  });
  pgRows.forEach((q, i) => kunci.appendRow([zone, i + 1, q.topic, q.question, q.correct ? 'BENAR' : 'SALAH']));
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
}

function getQuestionBanks() {
  return {
    1: {
      title:'UJIAN 1 • PENGUASAAN KONSEP',
      objective:[
        {type:'single',topic:'Konsep viskositas',question:'Dua cairan diuji pada kondisi yang sama. Cairan A mengalir lebih lambat daripada B. Kesimpulan yang paling tepat adalah...',options:['Viskositas A lebih besar','Massa jenis A pasti lebih besar','B pasti lebih panas','Gravitasi A lebih kecil'],answer:0,points:10,review:'Pahami makna viskositas sebagai hambatan internal terhadap aliran.'},
        {type:'single',topic:'Pengaruh suhu',question:'Untuk sebagian besar zat cair, kenaikan suhu menyebabkan...',options:['viskositas meningkat','viskositas menurun','viskositas selalu tetap','cairan berhenti mengalir'],answer:1,points:10,review:'Pelajari kembali hubungan suhu dengan viskositas zat cair.'},
        {type:'multi',topic:'Faktor viskositas',question:'Pilih dua pernyataan yang sesuai tentang viskositas zat cair.',options:['Viskositas lebih besar berarti aliran lebih terhambat','Kenaikan suhu umumnya menurunkan viskositas zat cair','Viskositas hanya ditentukan oleh warna cairan','Viskositas tidak berkaitan dengan gerak relatif lapisan fluida'],answer:[0,1],points:10,review:'Pelajari faktor suhu dan makna hambatan internal fluida.'},
        {type:'single',topic:'Hukum Newton tentang viskositas',question:'Dalam τ = η(dv/dy), besaran η menyatakan...',options:['massa jenis','viskositas dinamis','kecepatan terminal','tekanan'],answer:1,points:10,review:'Pelajari kembali simbol, makna, dan satuan viskositas dinamis.'},
        {type:'single',topic:'Air-minyak-madu',question:'Dalam pengamatan ViskoLab, urutan viskositas yang digunakan untuk membandingkan air, minyak, dan madu adalah...',options:['air < minyak < madu','madu < minyak < air','minyak < air < madu','air < madu < minyak'],answer:0,points:10,review:'Bandingkan kembali perilaku aliran air, minyak, dan madu pada Amati/Pahami.'}
      ],
      essay:[
        {topic:'Penjelasan konsep',question:'Jelaskan dengan bahasamu sendiri apa yang dimaksud dengan viskositas dan mengapa viskositas memengaruhi kemampuan suatu cairan mengalir.',points:10,rubric:'Menjelaskan viskositas sebagai ukuran hambatan internal terhadap aliran dan menghubungkannya dengan kemudahan/sulitnya cairan mengalir.',review:'Ulangi bagian Apa itu Viskositas? di Pahami.'},
        {topic:'Perbandingan fluida',question:'Air dan madu dituangkan dari wadah yang sama. Jelaskan mengapa madu mengalir lebih lambat.',points:10,rubric:'Menghubungkan aliran lebih lambat dengan viskositas madu yang lebih besar.',review:'Ulangi perbandingan Air–Minyak–Madu.'},
        {topic:'Suhu',question:'Jelaskan apa yang terjadi pada viskositas madu ketika madu dipanaskan dan bagaimana pengaruhnya terhadap aliran.',points:10,rubric:'Menjelaskan viskositas madu umumnya menurun saat suhu naik sehingga madu lebih mudah mengalir.',review:'Ulangi bagian pengaruh suhu terhadap viskositas.'},
        {topic:'Persamaan viskositas',question:'Jelaskan arti η, τ, dan dv/dy pada persamaan τ = η(dv/dy).',points:10,rubric:'Menjelaskan η sebagai viskositas dinamis, τ sebagai tegangan geser, dan dv/dy sebagai gradien kecepatan.',review:'Ulangi bagian Hukum Newton tentang viskositas dan rumusnya.'},
        {topic:'Fenomena sehari-hari',question:'Berikan satu contoh dalam kehidupan sehari-hari yang menunjukkan perbedaan viskositas dan jelaskan bukti pengamatannya.',points:10,rubric:'Contoh relevan dan penjelasan mengaitkan perbedaan kecepatan/kesulitan mengalir dengan viskositas.',review:'Ulangi contoh etnosains dan fenomena Amati.'}
      ]
    },
    2: {
      title:'UJIAN 2 • ANALISIS DATA DAN EKSPERIMEN',
      objective:[
        {type:'single',topic:'Analisis waktu alir',question:'Pada jarak yang sama, sampel X mengalir dalam 2 s dan Y dalam 8 s pada kondisi setara. Kesimpulan paling tepat adalah...',options:['X pasti lebih kental','Y cenderung memiliki viskositas lebih besar','Keduanya pasti memiliki viskositas sama','Waktu tidak dapat digunakan untuk membandingkan aliran'],answer:1,points:10,review:'Pelajari cara menghubungkan waktu pengamatan dengan hambatan aliran.'},
        {type:'single',topic:'Data suhu-viskositas',question:'Data menunjukkan η = 1,8 mPa·s pada 20°C dan η = 1,2 mPa·s pada 40°C. Kesimpulannya...',options:['η meningkat saat suhu naik','η menurun saat suhu naik','η tidak berubah','data tidak menunjukkan hubungan'],answer:1,points:10,review:'Baca kembali tabel/grafik hubungan suhu dan viskositas.'},
        {type:'multi',topic:'Hukum Stokes',question:'Jika F = 6πηrv, pilih dua perubahan yang membuat gaya hambat meningkat ketika besaran lain tetap.',options:['η diperbesar','r diperbesar','v diperkecil','η dibuat nol'],answer:[0,1],points:10,review:'Pelajari hubungan gaya hambat Stokes dengan η dan r.'},
        {type:'single',topic:'Kecepatan terminal',question:'Dari vₜ = 2r²g(ρs−ρf)/(9η), jika η diperbesar sementara besaran lain tetap, vₜ akan...',options:['meningkat','menurun','tetap','menjadi tak terhingga'],answer:1,points:10,review:'Pelajari kembali hubungan viskositas dan kecepatan terminal.'},
        {type:'single',topic:'Kualitas eksperimen',question:'Untuk menguji pengaruh suhu terhadap viskositas, variabel yang harus dikontrol antara percobaan adalah...',options:['jenis dan jumlah fluida serta kondisi alat','semua variabel diubah sekaligus','warna cairan saja','hasil akhir saja'],answer:0,points:10,review:'Ulangi konsep variabel kontrol dan desain eksperimen pada Eksplorasi.'}
      ],
      essay:[
        {topic:'Analisis data',question:'Suatu cairan menempuh 0,60 m dalam 3,0 s pada kondisi terminal. Hitung kecepatan terminal dan jelaskan maknanya.',points:10,rubric:'Menghitung v = s/t = 0,20 m/s dan menjelaskan bahwa ini adalah kecepatan gerak terminal pada kondisi pengamatan.',review:'Ulangi perhitungan kecepatan terminal dari jarak dan waktu.'},
        {topic:'Interpretasi grafik',question:'Grafik viskositas terhadap suhu menurun dari kiri ke kanan. Jelaskan hubungan yang ditunjukkan grafik dan dampaknya pada aliran.',points:10,rubric:'Menjelaskan η menurun ketika suhu naik dan aliran cenderung lebih mudah/cepat.',review:'Ulangi analisis grafik suhu-viskositas.'},
        {topic:'Stokes',question:'Jelaskan bagaimana perubahan viskositas memengaruhi gaya hambat pada bola yang bergerak dalam fluida menurut hukum Stokes.',points:10,rubric:'Menggunakan F = 6πηrv dan menjelaskan bahwa jika η naik, gaya hambat naik ketika r dan v tetap.',review:'Ulangi hukum Stokes.'},
        {topic:'Kecepatan terminal',question:'Dua bola identik dijatuhkan pada dua fluida. Bola A memiliki kecepatan terminal lebih kecil. Jelaskan kemungkinan penyebabnya.',points:10,rubric:'Menghubungkan v_t yang lebih kecil dengan viskositas fluida yang lebih besar jika parameter lain sama.',review:'Ulangi hubungan η dan v_t.'},
        {topic:'Evaluasi eksperimen',question:'Rancang secara singkat langkah eksperimen untuk membandingkan pengaruh suhu terhadap viskositas dua sampel cairan.',points:10,rubric:'Memuat variabel suhu sebagai variabel bebas, jenis/jumlah fluida dan alat dikontrol, pengukuran konsisten, pencatatan data, dan perbandingan hasil.',review:'Ulangi bagian prosedur eksperimen dan variabel kontrol.'}
      ]
    },
    3: {
      title:'UJIAN 3 • ANALISIS DAN PENERAPAN',
      objective:[
        {type:'single',topic:'Poiseuille',question:'Dalam hukum Poiseuille, laju aliran sangat sensitif terhadap perubahan...',options:['jari-jari pipa karena bergantung pada r⁴','warna pipa','massa pipa','bentuk luar wadah'],answer:0,points:10,review:'Pelajari kembali hubungan laju aliran dan jari-jari pipa.'},
        {type:'single',topic:'Aplikasi industri',question:'Pabrik harus mengalirkan cairan sangat kental. Pertimbangan fisika yang paling relevan adalah...',options:['tekanan, kondisi suhu, dan hambatan aliran','warna cairan saja','mengabaikan viskositas','menganggap semua fluida seperti air'],answer:0,points:10,review:'Ulangi penerapan viskositas pada sistem industri.'},
        {type:'multi',topic:'Analisis grafik',question:'Sebuah grafik menunjukkan suhu naik dan viskositas turun. Pilih dua kesimpulan yang sesuai.',options:['Cairan cenderung lebih mudah mengalir','Hambatan viskos internal cenderung berkurang','Cairan pasti berubah menjadi gas','Massa jenis pasti menjadi nol'],answer:[0,1],points:10,review:'Ulangi interpretasi grafik dan makna fisik penurunan viskositas.'},
        {type:'single',topic:'Lingkungan',question:'Jika terjadi tumpahan cairan sangat kental, dibandingkan cairan kurang kental pada kondisi sebanding, cairan tersebut cenderung...',options:['menyebar lebih lambat','selalu menyebar lebih cepat','tidak dipengaruhi permukaan','langsung menguap'],answer:0,points:10,review:'Ulangi penerapan konsep viskositas pada fenomena lingkungan.'},
        {type:'single',topic:'Pemecahan masalah',question:'Sistem pemompaan mengalami penurunan laju aliran setelah viskositas cairan meningkat. Tindakan analitis yang paling masuk akal adalah...',options:['mengevaluasi kondisi suhu, tekanan, dan karakteristik pipa','menghapus pengukuran viskositas','menganggap pompa pasti rusak','mengabaikan diameter pipa'],answer:0,points:10,review:'Ulangi penerapan viskositas pada aliran dalam pipa.'}
      ],
      essay:[
        {topic:'Poiseuille',question:'Jelaskan mengapa perubahan kecil jari-jari pipa dapat memberikan perubahan besar pada laju aliran dalam hukum Poiseuille.',points:10,rubric:'Menjelaskan bahwa laju aliran bergantung kuat pada r^4 sehingga perubahan jari-jari diperkuat oleh pangkat empat.',review:'Ulangi hukum Poiseuille.'},
        {topic:'Kasus industri',question:'Sebuah industri menggunakan cairan yang menjadi lebih kental saat suhu turun. Jelaskan strategi fisika yang dapat dipertimbangkan agar aliran tetap efektif.',points:10,rubric:'Menghubungkan kenaikan suhu dengan penurunan viskositas zat cair serta mempertimbangkan tekanan/pompa dan kondisi pipa.',review:'Ulangi penerapan suhu dan viskositas pada industri.'},
        {topic:'Kasus lingkungan',question:'Jelaskan mengapa viskositas perlu dipertimbangkan ketika menganalisis penyebaran suatu cairan di lingkungan.',points:10,rubric:'Menghubungkan viskositas dengan hambatan aliran dan kecenderungan penyebaran/gerak cairan.',review:'Ulangi aplikasi viskositas pada lingkungan.'},
        {topic:'Analisis eksperimen',question:'Data percobaan menunjukkan dua sampel memiliki waktu alir berbeda. Jelaskan bagaimana kamu menentukan sampel mana yang lebih viskos tanpa mengubah variabel lain.',points:10,rubric:'Menjelaskan kontrol kondisi, jarak sama, prosedur sama, pengukuran waktu, lalu menghubungkan waktu alir lebih besar dengan viskositas lebih tinggi secara konseptual.',review:'Ulangi analisis data eksperimen.'},
        {topic:'Sintesis konsep',question:'Hubungkan viskositas, gaya hambat, dan kecepatan terminal dalam satu penjelasan yang runtut.',points:10,rubric:'Menjelaskan viskositas lebih besar meningkatkan hambatan, sehingga pada kondisi Stokes kecepatan terminal cenderung lebih kecil.',review:'Ulangi hubungan viskositas–Stokes–kecepatan terminal.'}
      ]
    }
  };
}
