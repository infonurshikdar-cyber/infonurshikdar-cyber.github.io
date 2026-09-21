লাইভ সাইটে Photo System চালু করার জন্য সংক্ষিপ্ত ধাপ:

1) GitHub-এর cricket-stats repository-তে এই ZIP খুলে index.html এবং 404.html আপলোড/replace করুন।
2) Apps Script project-এ Code.gs-এর জায়গায় এই ZIP-এর Code.gs বসান।
3) Apps Script: Deploy > Manage deployments > Web app > Edit > New version > Deploy করুন।
4) Web app URL আগের একই URL রাখুন।
5) GitHub Pages কয়েক মুহূর্ত পরে নতুন index.html দেখাবে।
6) Admin PIN দিয়ে ঢুকে নতুন Player যোগ করার সময় Photo নির্বাচন করুন।
7) পুরোনো Player-এর ছবি বদলাতে Manage Player থেকে Player নির্বাচন > Photo নির্বাচন > "ছবি সেভ করুন"।

নোট: GitHub Pages শুধু frontend; Photo Save/Read Apps Script backend-এর মাধ্যমে হবে।
