import os
import sys

import torch
from torchvision import transforms
from torchvision.models import mobilenet_v3_small
from PIL import Image, ImageOps

# 이 파일과 같이 배포되는 사전학습 가중치(모델 자체는 ImageNet으로 학습된 것을
# 그대로 씀, 별도 학습 없음)를 읽어서 이미지 특징을 뽑는 용도로만 사용.
# 인터넷 연결 없이 완전히 이 PC에서 동작함(API 아님).
def _resource_path(*parts):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


MODEL_WEIGHTS_PATH = _resource_path("models", "mobilenet_v3_small-047dcff4.pth")

_PREPROCESS = transforms.Compose([
    transforms.Resize(256),
    transforms.CenterCrop(224),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225]),
])

_cos = torch.nn.CosineSimilarity(dim=0)
_model = None


def _get_model():
    """모델은 처음 쓸 때 한 번만 읽어서 메모리에 올려두고 계속 재사용."""
    global _model
    if _model is None:
        model = mobilenet_v3_small(weights=None)
        state_dict = torch.load(MODEL_WEIGHTS_PATH, map_location="cpu")
        model.load_state_dict(state_dict)
        model.classifier = torch.nn.Identity()  # 분류기는 떼고 특징(임베딩)만 뽑음
        model.eval()
        _model = model
    return _model


def _embed(path):
    model = _get_model()
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    x = _PREPROCESS(im).unsqueeze(0)
    with torch.no_grad():
        feat = model(x)
    return feat.squeeze(0)


def classify_photos(photo_paths, reference_photos, log=print):
    """사전학습된 이미지 인식 모델(MobileNetV3-Small)로 사진의 특징을 뽑아,
    단계별 기준 사진들과 가장 비슷한 단계로 분류함.
    reference_photos: {단계이름: [기준사진경로, ...]} (단계마다 여러 장일수록 정확도가 좋아짐)
    반환값: {사진경로: 단계이름}"""
    ref_embeds = {}
    for name, paths in reference_photos.items():
        embeds = []
        for p in paths:
            try:
                embeds.append(_embed(p))
            except Exception as e:
                log(f"  기준 사진 '{name}' - {os.path.basename(p)} 읽기 실패: {e}")
        if embeds:
            ref_embeds[name] = embeds

    if not ref_embeds:
        raise RuntimeError("기준 사진을 하나도 읽지 못했습니다.")

    assignments = {}
    total = len(photo_paths)
    for i, p in enumerate(photo_paths, 1):
        try:
            emb = _embed(p)
        except Exception as e:
            log(f"  [{i}/{total}] {os.path.basename(p)} 읽기 실패: {e}")
            continue

        best_name, best_sim = None, None
        for name, embeds in ref_embeds.items():
            sim = max(_cos(emb, re).item() for re in embeds)
            if best_sim is None or sim > best_sim:
                best_name, best_sim = name, sim

        log(f"  [{i}/{total}] {os.path.basename(p)} -> {best_name} (유사도 {best_sim:.3f})")
        assignments[p] = best_name

    return assignments
