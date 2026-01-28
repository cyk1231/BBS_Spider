import os
import time
import json
from urllib.parse import unquote
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# --- 配置 ---
OUT_DIR = "WuShu_Archive/images_all"
PROGRESS_FILE = "download_progress.json"
os.makedirs(OUT_DIR, exist_ok=True)

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r') as f:
            return set(json.load(f))
    return set()

def save_progress(seen_hrefs):
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(list(seen_hrefs), f)

def pick_best_frame(page_or_context):
    """自动寻找包含附件链接的框架"""
    best = None
    best_score = -1
    # 遍历当前页面及其所有子框架
    for fr in page_or_context.frames:
        try:
            dn = fr.locator("a[href*='dn?']").count()
            sf = fr.locator("a[href*='sf?']").count()
            score = dn * 10 + sf
            if score > best_score:
                best_score = score
                best = fr
        except:
            pass
    return best

def download_attachments(post_page, out_dir, post_title):
    """只下载 dn 附件，并按：帖子标题_原文件名 保存"""
    fr = pick_best_frame(post_page)
    if not fr:
        return False

    ctx = post_page.context
    dn_links = fr.locator("a[href*='dn?']")
    n = dn_links.count()
    if n == 0:
        return False

    success = False
    title_part = safe_filename(post_title, max_len=60)  # 标题别太长，避免路径过长

    for i in range(n):
        link = dn_links.nth(i)
        before = set(ctx.pages)

        try:
            # 尽量减少新开页
            try:
                link.evaluate("el => el.removeAttribute('target')")
            except:
                pass

            with post_page.expect_download(timeout=10000) as dlinfo:
                link.click(no_wait_after=True)

            dl = dlinfo.value
            fname = unquote(dl.suggested_filename)
            fname_part = safe_filename(fname, max_len=120)

            # ✅ 核心：帖子标题_图片本身名字
            final_name = f"{title_part}_{fname_part}"
            save_path = unique_path(out_dir, final_name)

            dl.save_as(save_path)
            print(f"    [√] 下载成功: {final_name}")
            success = True

        except PWTimeout:
            print("    [!] dn 下载等待超时，跳过该附件")
        except Exception as e:
            print(f"    [!] 下载失败，跳过: {e}")
        finally:
            # 关闭本次点击新弹出的预览页
            for p in list(ctx.pages):
                if p not in before and p != post_page:
                    try:
                        p.close()
                    except:
                        pass

    return success


def safe_filename(s, max_len=80):
    # Windows 不允许的字符过滤 + 控制长度
    s = "".join(c for c in s if c not in r'\/:*?"<>|').strip()
    return s[:max_len] if len(s) > max_len else s

def unique_path(dir_path, base_name):
    """
    防止同名覆盖：若已存在，则自动追加 _1 _2 ...
    base_name: 不含路径的文件名
    """
    path = os.path.join(dir_path, base_name)
    if not os.path.exists(path):
        return path
    name, ext = os.path.splitext(base_name)
    k = 1
    while True:
        path2 = os.path.join(dir_path, f"{name}_{k}{ext}")
        if not os.path.exists(path2):
            return path2
        k += 1


def run():
    seen_posts = load_history()
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto("https://wvpn.ustc.edu.cn/")
        print(">>> 请手动登录 WebVPN 并进入【精武门】板块列表页...")
        input(">>> 到达第一页列表后，按回车开始...")

        while True:
            # 1. 切换到列表所在的框架 (通常是 f3)
            list_frame = page.frame_locator('frame[name="f3"]')
            
            try:
                list_frame.locator("a.o_title").first.wait_for(timeout=10000)
                post_links = list_frame.locator("a.o_title").all()
            except:
                print(">>> 未发现帖子列表，任务结束。")
                break

            print(f"\n>>> 正在扫描本页 {len(post_links)} 个帖子...")

            for link in post_links:
                href = link.get_attribute("href")
                title = link.inner_text().strip()

                if href in seen_posts:
                    continue

                try:
                    # 在新标签页打开帖子，避免列表页框架刷新丢失
                    with context.expect_page() as new_page_info:
                        link.click(modifiers=["Control"])
                    post_page = new_page_info.value
                    post_page.wait_for_load_state("domcontentloaded")

                    # 执行你测试成功的下载逻辑
                    has_img = download_attachments(post_page, OUT_DIR, title)
                    if has_img:
                        print(f"  [发现附件] {title}")

                    # ✅ 本帖附件都下载完后：关闭期间新打开的 sf/dn 预览标签页
                    close_extra_pages_after_post(context, keep_pages=[page, post_page])

                    post_page.close()
                    seen_posts.add(href)

                except Exception as e:
                    print(f"  [!] 处理失败 {title}: {e}")

            # 进度保存
            save_progress(seen_posts)

            # 2. 翻页逻辑
            next_btn = list_frame.locator("a.next").first
            is_disabled = "disabled" in (next_btn.get_attribute("class") or "")
            if next_btn.is_visible() and not is_disabled:
                print(">>> 翻向下一页...")
                next_btn.click()
                time.sleep(3) # 给 WebVPN 留一点重写 URL 的时间
            else:
                break

        print("\n>>> 全部下载任务完成！")
        browser.close()

def load_history():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r') as f:
            return set(json.load(f))
    return set()

def close_extra_pages_after_post(context, keep_pages):
    """
    关闭本帖下载过程中弹出来的多余标签页（如 sf 图片预览页）。
    keep_pages: 需要保留的 Page 列表，例如 [主列表页 page, 当前帖子页 post_page]
    """
    keep = {p for p in keep_pages if p is not None}
    for p in list(context.pages):
        if p not in keep:
            try:
                p.close()
            except:
                pass


if __name__ == "__main__":
    run()