"""Generate SumSelect app icon candidates through the local ComfyUI API."""
import json, time, urllib.request, uuid, os, sys

HOST = "http://127.0.0.1:8188"
OUT = r"C:\SumSelect\icons"
os.makedirs(OUT, exist_ok=True)

POS = ("app icon, flat vector illustration, rounded square tile, bold greek letter sigma glyph centered, "
       "mint green and teal gradient, soft inner shadow, crisp geometric edges, minimal, "
       "clean dark charcoal background, centered composition, high contrast, no text, no letters except sigma")
NEG = "photo, realistic, 3d render, text, words, watermark, signature, blurry, noisy, clutter, busy background"


def workflow(seed, pos=POS):
    return {
        "1": {"class_type": "UNETLoader",
              "inputs": {"unet_name": "qwen_image_2.1_int8_convrot.safetensors", "weight_dtype": "default"}},
        "2": {"class_type": "CLIPLoader",
              "inputs": {"clip_name": "qwen3vl_8b_w4a8.safetensors", "type": "qwen_image"}},
        "3": {"class_type": "VAELoader", "inputs": {"vae_name": "qwen_image_2.1_vae_bf16.safetensors"}},
        "4": {"class_type": "CLIPTextEncode", "inputs": {"text": pos, "clip": ["2", 0]}},
        "5": {"class_type": "CLIPTextEncode", "inputs": {"text": NEG, "clip": ["2", 0]}},
        "6": {"class_type": "EmptySD3LatentImage", "inputs": {"width": 1024, "height": 1024, "batch_size": 1}},
        "7": {"class_type": "ModelSamplingAuraFlow", "inputs": {"model": ["1", 0], "shift": 3.1}},
        "8": {"class_type": "KSampler",
              "inputs": {"model": ["7", 0], "seed": seed, "steps": 20, "cfg": 2.5,
                         "sampler_name": "euler", "scheduler": "simple", "denoise": 1.0,
                         "positive": ["4", 0], "negative": ["5", 0], "latent_image": ["6", 0]}},
        "9": {"class_type": "VAEDecode", "inputs": {"samples": ["8", 0], "vae": ["3", 0]}},
        "10": {"class_type": "SaveImage", "inputs": {"images": ["9", 0], "filename_prefix": "sumselect_icon"}},
    }


def post(wf):
    body = json.dumps({"prompt": wf, "client_id": str(uuid.uuid4())}).encode()
    req = urllib.request.Request(HOST + "/prompt", body, {"Content-Type": "application/json"})
    return json.load(urllib.request.urlopen(req, timeout=30))["prompt_id"]


def wait(pid, timeout=900):
    t0 = time.time()
    while time.time() - t0 < timeout:
        h = json.load(urllib.request.urlopen(f"{HOST}/history/{pid}", timeout=30))
        if pid in h:
            entry = h[pid]
            status = entry.get("status", {})
            if status.get("status_str") == "error" or not status.get("completed", True):
                msgs = json.dumps(status.get("messages", []))[:800]
                if "error" in msgs.lower():
                    return None, msgs
            outs = entry.get("outputs", {})
            imgs = [i for o in outs.values() for i in o.get("images", [])]
            if imgs:
                return imgs, None
            if status.get("completed"):
                return [], json.dumps(status)[:500]
        time.sleep(3)
    return None, "timeout"


def fetch(img, dest):
    q = urllib.parse.urlencode({"filename": img["filename"], "subfolder": img.get("subfolder", ""),
                                "type": img.get("type", "output")})
    data = urllib.request.urlopen(f"{HOST}/view?{q}", timeout=60).read()
    open(dest, "wb").write(data)
    return dest


if __name__ == "__main__":
    seeds = [int(x) for x in sys.argv[1:]] or [7]
    for seed in seeds:
        pid = post(workflow(seed))
        print("queued", seed, pid, flush=True)
        imgs, err = wait(pid)
        if err:
            print("ERROR", seed, err, flush=True)
            continue
        for n, img in enumerate(imgs):
            path = os.path.join(OUT, f"icon_{seed}_{n}.png")
            fetch(img, path)
            print("saved", path, flush=True)
