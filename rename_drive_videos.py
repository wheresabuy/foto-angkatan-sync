#!/usr/bin/env python3
"""
Rename Video Files in Buku Angkatan (Google Drive & Local)
Format: [3 Digit Terakhir NRP]_[Nama Mahasiswa].[ext]
Contoh: 001_Nabatan Hasana.mp4
"""

import os
import re
import json
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor

ENV_FILE = "/home/abuyyy/Foto_Teman_3x4_A4/.env.github"
LOCAL_ORGANIZED_DIR = "/home/abuyyy/Buku_Angkatan_Organized"
DRIVE_ROOT_FOLDER_ID = "1HS4EInIOlsWEXKPP05sSapKE1Dt0G72W"

DEFAULT_CLIENT_ID = "44438659992-7kgjeitenc16ssihbtdjbgguch7ju55s.apps.googleusercontent.com"
DEFAULT_CLIENT_SECRET = "-gMLuQyDiI0XrQS_vx_mhuYF"

def get_access_token():
    refresh_token = os.environ.get("GDRIVE_REFRESH_TOKEN")
    client_id = os.environ.get("GDRIVE_CLIENT_ID", DEFAULT_CLIENT_ID).strip().strip('"\'')
    client_secret = os.environ.get("GDRIVE_CLIENT_SECRET", DEFAULT_CLIENT_SECRET).strip().strip('"\'')

    if not refresh_token and os.path.exists(ENV_FILE):
        with open(ENV_FILE) as f:
            for line in f:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    if k == "GDRIVE_REFRESH_TOKEN":
                        refresh_token = v.strip().strip('"\'')
                    elif k == "GDRIVE_CLIENT_ID":
                        client_id = v.strip().strip('"\'')
                    elif k == "GDRIVE_CLIENT_SECRET":
                        client_secret = v.strip().strip('"\'')

    if not refresh_token:
        raise ValueError("Refresh token not found!")

    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token"
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"}
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode("utf-8"))["access_token"]

def list_all(token, q, fields):
    files = []
    page_token = None
    while True:
        url = f"https://www.googleapis.com/drive/v3/files?q={urllib.parse.quote(q)}&fields=nextPageToken,{fields}&pageSize=100"
        if page_token:
            url += f"&pageToken={page_token}"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req) as resp:
            res = json.loads(resp.read().decode("utf-8"))
            files.extend(res.get("files", []))
            page_token = res.get("nextPageToken")
            if not page_token:
                break
    return files

def rename_drive_file(token, file_id, new_name):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
    body = json.dumps({"name": new_name}).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        method="PATCH"
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode("utf-8"))

def parse_folder_name(folder_name):
    m = re.match(r"^(\d{3})\s*-\s*(.+)$", folder_name)
    if m:
        nrp_3 = m.group(1)
        nama = m.group(2).strip().rstrip(".")
        return nrp_3, nama
    return None, None

def run_drive_renaming(token):
    print("\n" + "="*60)
    print("MENGAMBIL DATA DARI GOOGLE DRIVE...")
    print("="*60)

    # 1. Ambil seluruh subfolder di bawah folder Buku Angkatan
    subfolders = list_all(
        token,
        f"'{DRIVE_ROOT_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        "files(id,name)"
    )
    folder_by_id = {f["id"]: f["name"] for f in subfolders}
    print(f"Ditemukan {len(folder_by_id)} subfolder mahasiswa di Google Drive.")

    # 2. Ambil seluruh file video di Drive
    all_videos = list_all(token, "mimeType contains 'video/' and trashed = false", "files(id,name,mimeType,parents)")
    
    # 3. Petakan video ke folder mahasiswa
    student_videos = []
    for v in all_videos:
        for p in v.get("parents", []):
            if p in folder_by_id:
                student_videos.append((folder_by_id[p], v["id"], v["name"]))
                break

    student_videos.sort(key=lambda x: x[0])
    print(f"Ditemukan {len(student_videos)} file video di dalam folder mahasiswa.")

    # 4. Buat rencana rename
    by_folder = {}
    for folder_name, vid_id, vid_name in student_videos:
        by_folder.setdefault(folder_name, []).append((vid_id, vid_name))

    rename_tasks = []
    for folder_name, vids in by_folder.items():
        nrp_3, nama = parse_folder_name(folder_name)
        if not nrp_3 or not nama:
            continue
        for idx, (vid_id, vid_name) in enumerate(vids):
            ext = os.path.splitext(vid_name)[1] or ".mp4"
            if len(vids) == 1:
                target_name = f"{nrp_3}_{nama}{ext}"
            else:
                target_name = f"{nrp_3}_{nama}_{idx+1}{ext}"

            if vid_name != target_name:
                rename_tasks.append((folder_name, vid_id, vid_name, target_name))
            else:
                print(f"[SUDAH SESUAI] {folder_name} -> {vid_name}")

    print(f"\nTotal video yang perlu di-rename di Google Drive: {len(rename_tasks)}")
    
    # 5. Eksekusi rename secara paralel (10 worker)
    success_count = 0
    fail_count = 0

    def do_rename(item):
        folder_name, vid_id, old_name, new_name = item
        try:
            rename_drive_file(token, vid_id, new_name)
            return True, f"[{folder_name}] {old_name} -> {new_name}"
        except Exception as e:
            return False, f"[GAGAL] [{folder_name}] {old_name}: {e}"

    with ThreadPoolExecutor(max_workers=10) as executor:
        for ok, msg in executor.map(do_rename, rename_tasks):
            if ok:
                success_count += 1
                print(f"[OK] {msg}")
            else:
                fail_count += 1
                print(msg)

    print(f"\nRingkasan Google Drive: {success_count} berhasil di-rename, {fail_count} gagal.")

def run_local_renaming():
    if not os.path.exists(LOCAL_ORGANIZED_DIR):
        print(f"\nFolder lokal {LOCAL_ORGANIZED_DIR} tidak ditemukan, lewati.")
        return

    print("\n" + "="*60)
    print("MEMERIKSA FOLDER LOKAL: Buku_Angkatan_Organized...")
    print("="*60)

    video_exts = {".mp4", ".mov", ".mkv", ".avi", ".3gp"}
    local_renames = 0

    for folder in sorted(os.listdir(LOCAL_ORGANIZED_DIR)):
        folder_path = os.path.join(LOCAL_ORGANIZED_DIR, folder)
        if not os.path.isdir(folder_path):
            continue
        nrp_3, nama = parse_folder_name(folder)
        if not nrp_3 or not nama:
            continue

        files = os.listdir(folder_path)
        vids = [f for f in files if os.path.splitext(f)[1].lower() in video_exts]
        vids.sort()

        for idx, vid_name in enumerate(vids):
            ext = os.path.splitext(vid_name)[1]
            if len(vids) == 1:
                target_name = f"{nrp_3}_{nama}{ext}"
            else:
                target_name = f"{nrp_3}_{nama}_{idx+1}{ext}"

            if vid_name != target_name:
                old_path = os.path.join(folder_path, vid_name)
                new_path = os.path.join(folder_path, target_name)
                os.rename(old_path, new_path)
                print(f"[LOCAL OK] [{folder}] {vid_name} -> {target_name}")
                local_renames += 1
            else:
                print(f"[LOCAL SUDAH SESUAI] [{folder}] {vid_name}")

    print(f"Selesai! Total video di folder lokal yang di-rename: {local_renames}")

def main():
    print("Mendapatkan Access Token Google Drive...")
    token = get_access_token()
    print("Token berhasil diperoleh.")

    # 1. Rename di Google Drive
    run_drive_renaming(token)

    # 2. Rename di folder lokal jika ada
    run_local_renaming()

    print("\nSemua proses penggantian nama video selesai!")

if __name__ == "__main__":
    main()
