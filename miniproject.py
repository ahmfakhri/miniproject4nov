import pandas as pd
import matplotlib.pyplot as plt

file_path = "Data_Wisudawan.xlsx"
data = pd.read_excel(file_path)


#hitung jumlah wisudawan per prodi pake groupby
jumlahPerProdi = data.groupby('Program Studi')['NIM'].size().reset_index(name='Jumlah Wisudawan')

def klasifikasigrade(ipk):
    if ipk >= 3.75:
        return 'A'
    elif ipk >= 3.50:
        return 'B+'
    elif ipk >= 3.00:
        return 'B'
    elif ipk >= 2.50:
        return 'C'
    else:
        return 'D'
    
def klasifikasipredikat(ipk):
    if ipk >= 3.75:
        return 'Cumlaude'
    elif ipk >= 3.50:
        return 'Sangat Memuaskan'
    elif ipk >= 3.00:
        return 'Memuaskan'
    else:
        return 'Cukup'

#apply fungsinya ke kolom IPK, masukin ke Grade dan Predikat
data['Grade'] = data['IPK'].apply(klasifikasigrade)
data['Predikat'] = data['IPK'].apply(klasifikasipredikat)

#menyusun urutan kolom utuk output
data = data[["NIM","Nama Mahasiswa", "Program Studi", "IPK", "Lama Studi (Semester)", "Grade", "Predikat", "Tahun Wisuda"]]

print(data.head())
print(jumlahPerProdi)

# Simpan hasil ke file baru
data.to_excel("Rekap_Wisuda_Final.xlsx", index=False)

#grafik jumlah wisudawan per prodi
plt.figure(figsize=(8,5))
plt.bar(jumlahPerProdi['Program Studi'], jumlahPerProdi['Jumlah Wisudawan'])
plt.title('Jumlah Wisudawan per Program Studi')
plt.xlabel('Program Studi')
plt.ylabel('Jumlah Wisudawan')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()


#pie chart predikat
predikatCounts = data['Predikat'].value_counts()
plt.figure(figsize=(8,8))
plt.pie(predikatCounts, labels=predikatCounts.index, autopct='%.1f%%', startangle=140)
plt.title('Distribusi Predikat Wisudawan')
plt.axis('equal')
plt.show()


#rata-rata ipk per prodi
rataIPK = data.groupby('Program Studi')['IPK'].mean()

plt.figure(figsize=(8, 5))
rataIPK.plot(kind='bar', color='lightpink', edgecolor='black')
plt.title('Rata-Rata IPK per Program Studi')
plt.xlabel('Program Studi')
plt.ylabel('Rata-Rata IPK')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()

#lama studi rata-rata per prodi
rataLamaStudi = data.groupby('Program Studi')['Lama Studi (Semester)'].mean()
plt.figure(figsize=(8, 5))
rataLamaStudi.plot(kind='bar', color='salmon', edgecolor='black')
plt.title('Rata-Rata Lama Studi per Program Studi')
plt.xlabel('Program Studi')
plt.ylabel('Rata-Rata Lama Studi (Semester)')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()

print(f"\nAnalisis selesai! Output saved in : Rekap_Wisuda_Final.xlsx")
