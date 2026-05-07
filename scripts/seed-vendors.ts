
import { initializeApp } from "firebase/app";
import { getFirestore, collection, addDoc, serverTimestamp } from "firebase/firestore";
import fs from 'fs';
import path from 'path';

// Read config
const configPath = path.resolve(process.cwd(), 'firebase-applet-config.json');
const firebaseConfig = JSON.parse(fs.readFileSync(configPath, 'utf-8'));

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const vendors = [
  { name: "(주)경동하이테크판매", phone: "031-797-6377", fax: "031-797-6379", representative: "김명애", businessNumber: "159-86-01270", email: "kdhitech@naver.com" },
  { name: "주식회사 그린씨엠", phone: "032-322-8708", fax: "032-322-8709", representative: "김승철", businessNumber: "130-86-48798", email: "greencm65@mail.com" },
  { name: "그린텍", phone: "010-3893-4218", fax: "02-908-4218", representative: "설춘현", businessNumber: "210-18-37880", email: "softseol1@hanmail.net" },
  { name: "(주)다온홈시스", phone: "031-387-0585", fax: "031-388-0586", representative: "임현범", businessNumber: "138-81-52840", email: "daonhomsys@naver.com" },
  { name: "(주)세경공조", phone: "031-334-6902", fax: "", representative: "성진수", businessNumber: "142-81-66358", email: "sunjeni74@naver.com" }
];

async function seed() {
  console.log("Seeding vendors...");
  for (const v of vendors) {
    try {
      await addDoc(collection(db, "vendors"), {
        ...v,
        password: "1234", // Default password for seeded vendors
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      console.log(`Added: ${v.name}`);
    } catch (e) {
      console.error(`Error adding ${v.name}:`, e);
    }
  }
  console.log("Seeding complete.");
  process.exit(0);
}

seed();
