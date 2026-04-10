import pandas as pd
import matplotlib.pyplot as plt

data = [
    ["01.11","Growing of cereals","0111","Growing of cereals","One-to-One","Same activity definition"],
    ["01.62","Animal support","0162","Animal support","One-to-One","Direct match"],
    ["08.11","Stone quarrying","0810","Stone, sand, clay","Approximate","ISIC is broader"],
    ["10.13","Meat processing","1010","Meat processing","Approximate","ISIC is broader"],
    ["13.10","Textile fibres","1311","Textile fibres","One-to-One","Same scope"],
    ["20.11","Industrial gases","2011","Basic chemicals","Approximate","ISIC groups categories"],
    ["21.10","Pharma basic","2100","Pharma products","Approximate","ISIC combines categories"],
    ["26.20","Computers manufacturing","2620","Computers manufacturing","One-to-One","Exact match"],
    ["27.11","Electric motors","2710","Electrical equipment","Approximate","ISIC broader"],
    ["29.10","Motor vehicles","2910","Motor vehicles","One-to-One","Exact match"],
    ["35.11","Electricity production","3510","Electricity supply","Approximate","ISIC combines activities"],
    ["36.00","Water supply","3600","Water supply","One-to-One","Exact match"],
    ["41.10","Building development","4100","Construction","Approximate","ISIC combines"],
    ["43.21","Electrical installation","4321","Electrical installation","One-to-One","Exact match"],
    ["46.51","Wholesale IT","4651","Wholesale IT","One-to-One","Exact match"],
    ["47.11","Retail stores","4711","Retail stores","One-to-One","Exact match"],
    ["49.10","Rail transport","4911","Rail transport","One-to-One","Exact match"],
    ["62.01","Programming","6201","Programming","One-to-One","Exact match"],
    ["62.02","IT consultancy","6202","IT services","One-to-Many","ISIC splits services"],
    ["63.11","Data hosting","6311","Data hosting","One-to-One","Exact match"],
    ["68.20","Real estate renting","6820","Real estate","One-to-One","Exact match"],
    ["70.22","Business consulting","7020","Consulting","Approximate","ISIC merges categories"],
    ["72.11","Biotech R&D","7210","Science R&D","Approximate","ISIC broader"],
    ["81.21","Cleaning buildings","8121","Cleaning","One-to-One","Exact match"],
    ["86.10","Hospitals","8610","Hospitals","One-to-One","Exact match"]
]

columns = ["NACE Code","NACE Description","ISIC Code","ISIC Description","Match Type","Rationale"]

df = pd.DataFrame(data, columns=columns)

summary = df["Match Type"].value_counts()

print("\nMatch Type Summary:\n")
print(summary)

plt.figure()
summary.plot(kind="bar")
plt.title("Match Type Distribution")
plt.xlabel("Match Type")
plt.ylabel("Count")
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig("match_type_bar.png")
plt.close()

plt.figure()
summary.plot(kind="pie", autopct='%1.0f%%')
plt.title("Match Type Share")
plt.ylabel("")
plt.tight_layout()
plt.savefig("match_type_pie.png")
plt.close()

file_name = "Patel_Axit_MappingTask_NACE_to_ISIC.xlsx"

with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Mapping", index=False)
    summary = df["Match Type"].value_counts().sort_index()
    summary_df = summary.to_frame(name="Count")
    summary_df["Percentage"] = (summary_df["Count"] / len(df)) * 100
    summary_df.to_excel(writer, sheet_name="Summary")

print("\nExcel file + charts created successfully.")