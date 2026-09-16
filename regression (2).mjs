import { classifyCategory, safeProductScore, normalizedPrice, eachPrice, measurement, nextCategoryRange } from "./procurement.js";
import fs from "fs";

// Load OUR full dictionary from the SQL we built and tested
const sql = fs.readFileSync("./restaurant_food_dictionary_v1.sql","utf8");
const rows=[...sql.matchAll(/\('Restaurant','([^']+)','(\[.*?\])'::jsonb,\d+\)/g)];
const cats = rows.map(m=>({name:m[1], keywords:JSON.parse(m[2])}));
console.log("Loaded dictionary categories:", cats.map(c=>`${c.name}(${c.keywords.length})`).join(", "));

let pass=0, fail=0;
const t=(label,got,exp)=>{const ok=String(got)===String(exp);console.log(`${ok?"PASS":"FAIL"}  ${label.padEnd(46)} got=${String(got).padEnd(14)}exp=${exp}`);ok?pass++:fail++;};

console.log("\n-- CLASSIFICATION (byproducts must not be meat/dairy/produce) --");
for(const [d,e] of [["CHICKEN BASE","General"],["BEEF BASE","General"],["CHICKEN BROTH","General"],
 ["COCONUT MILK","General"],["PEANUT BUTTER","General"],["IMITATION CRAB MEAT","General"],
 ["EGG SUBSTITUTE","General"],["POTATO CHIPS","General"],["ONION POWDER","General"],
 ["CHICKEN BREAST","Meat"],["BEEF TENDERLOIN","Meat"],["CRAB LEGS","Meat"],
 ["WHITE MUSHROOM","Produce"],["ROMA TOMATOES","Produce"],["EGGS X/LG","Dairy"],["UNSALTED BUTTER","Dairy"]])
  { const c=classifyCategory(d,cats); t(`"${d}"`, c?c.name:"UNCATEGORIZED", e); }

console.log("\n-- MATCHING (0.85 = auto-link threshold) --");
for(const [a,b,e] of [["Roma Tomatoes 25lb","Roma Tomatoes, 25 lb","LINK"],
 ["Chicken Breast 40lb","Chicken Thighs 40lb","review"],["Bacon 1lb","Bacon 5lb","review"],
 ["Tomato Sauce","Tomato Paste","review"],["TOMATO","GRAPE TOMATO","review"]])
  { const s=safeProductScore(a,b); t(`${a} | ${b}`, s>=0.85?"LINK":"review", e); }

console.log("\n-- UNITS (construction readiness) --");
t("8 ft is length", measurement(8,"ft")?.dimension, "length");
t("2 in is length", measurement(2,"in")?.dimension, "length");
t("50 lb is mass", measurement(50,"lb")?.dimension, "mass");
t("5 gal is volume", measurement(5,"gal")?.dimension, "volume");
t("24 ct is count", measurement(24,"ct")?.dimension, "count");
t("unknown unit kept", measurement(3,"sheet")?.dimension, "unknown");

console.log("\n-- PACK / PRICE --");
t("4/1 GAL each price", eachPrice(74.89,"4/1 GAL")?.price, 18.72);
t("50 LB normalized unit", normalizedPrice(25,"50 LB")?.unit, "G");
t("category block", JSON.stringify(nextCategoryRange([{range_end:19999}])), '{"range_start":20000,"range_end":29999}');

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail?1:0);
