import * as THREE from 'three';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { GLTFLoader } from 'three/addons/loaders/GLTFLoader.js';
import { DRACOLoader } from 'three/addons/loaders/DRACOLoader.js';
import { RGBELoader } from 'three/addons/loaders/RGBELoader.js';
import { EffectComposer } from 'three/addons/postprocessing/EffectComposer.js';
import { RenderPass } from 'three/addons/postprocessing/RenderPass.js';
import { OutputPass } from 'three/addons/postprocessing/OutputPass.js';
import { UnrealBloomPass } from 'three/addons/postprocessing/UnrealBloomPass.js';
import { ShaderPass } from 'three/addons/postprocessing/ShaderPass.js';
import { FXAAShader } from 'three/addons/shaders/FXAAShader.js';
import { SSAOPass } from 'three/addons/postprocessing/SSAOPass.js';
import { TransformControls } from 'three/addons/controls/TransformControls.js';
import { RectAreaLightHelper } from 'three/addons/helpers/RectAreaLightHelper.js';
import { RectAreaLightUniformsLib } from 'three/addons/lights/RectAreaLightUniformsLib.js';
import { computeBoundsTree, disposeBoundsTree, acceleratedRaycast } from 'three-mesh-bvh';
THREE.BufferGeometry.prototype.computeBoundsTree = computeBoundsTree;
THREE.BufferGeometry.prototype.disposeBoundsTree = disposeBoundsTree;
THREE.Mesh.prototype.raycast = acceleratedRaycast;
async function loadDataFromSheet() {
    // --- LOCAL EXCEL MODE ---
    try {
        const LOCAL_EXCEL_PATH = 'CO-EXCEL-MASTER-5.xlsx';
        const response = await fetch(LOCAL_EXCEL_PATH);
        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        const rows = rawData.map(row => {
            return {
                c: row.map(cell => {
                    return (cell !== null && cell !== undefined) ? { v: cell } : null;
                })
            };
        });

        sceneObjects = [];
        const getVal = (row, idx, type = 'string') => {
            const cell = row.c[idx];
            if (!cell || cell.v === null || cell.v === 'null') return type === 'number' ? 0 : "";
            return cell.v;
        };
        if (rows[0]) infoTextAbout = getVal(rows[0], 2);
        if (rows[2]) infoTextStatus = getVal(rows[2], 2);
        if (rows[3]) marqueeText = getVal(rows[3], 2);
        const _mobile = isMobileView();
        rows.forEach((row, i) => {
            let glbSrc = getVal(row, 10);
            if (_mobile) {
                const mobileGlb = getVal(row, 45);
                if (mobileGlb && typeof mobileGlb === 'string' && mobileGlb.trim() !== '' && mobileGlb.trim().toLowerCase() !== 'null' && (mobileGlb.endsWith('.glb') || mobileGlb.endsWith('.gltf'))) {
                    glbSrc = mobileGlb;
                }
            }
            if (!glbSrc || typeof glbSrc !== 'string' || glbSrc.trim() === '' || glbSrc.trim().toLowerCase() === 'null' || (!glbSrc.endsWith('.glb') && !glbSrc.endsWith('.gltf'))) return;
            const gridThumbRaw = getVal(row, 13);
            const gridThumbnail = (gridThumbRaw && typeof gridThumbRaw === 'string' && gridThumbRaw.trim() !== '' && gridThumbRaw.match(/\.(jpeg|jpg|gif|png|webp)$/i)) ? gridThumbRaw.trim() : '';
            const obj = {
                type: 'model',
                src: glbSrc,
                width: 700,
                rowIndex: i,
                gridThumbnail: gridThumbnail,
                thumbScale: (() => { const raw = parseFloat(getVal(row, 26, 'number')); return isNaN(raw) ? 1 : raw; })(),
                gridScale: parseFloat(getVal(row, 50, 'number')) || 0,
                gridRotX: parseFloat(getVal(row, 51, 'number')) || 0,
                gridRotY: parseFloat(getVal(row, 52, 'number')) || 0,
                gridOffsetY: parseFloat(getVal(row, 53, 'number')) || 0,
                position: {
                    x: parseFloat(getVal(row, 27, 'number')),
                    y: parseFloat(getVal(row, 28, 'number')),
                    z: parseFloat(getVal(row, 29, 'number'))
                },
                rotation: {
                    x: parseFloat(getVal(row, 30, 'number')),
                    y: parseFloat(getVal(row, 31, 'number')),
                    z: parseFloat(getVal(row, 32, 'number'))
                },
                scale: parseFloat(getVal(row, 33, 'number')) || 400,
                modalContent: {
                    artist: getVal(row, 2),
                    name: getVal(row, 3),
                    hoverDisplayName: getVal(row, 37),
                    price: getVal(row, 16),
                    dimensions: getVal(row, 15),
                    category: getVal(row, 5),
                    type: getVal(row, 6),
                    color: getVal(row, 7),
                    edition: getVal(row, 14),
                    description: getVal(row, 18),
                    bio: getVal(row, 22),
                    media: []
                }
            };
            const mainImg = getVal(row, 11);
            if (mainImg && mainImg.match(/\.(jpeg|jpg|gif|png|webp)$/i)) {
                obj.modalContent.media.push({ type: "image", src: mainImg });
            }
            const extraImgs = [34, 35, 36];
            extraImgs.forEach(idx => {
                const link = getVal(row, idx);
                if (link && typeof link === 'string' && link.trim() !== '') {
                    obj.modalContent.media.push({ type: "image", src: link });
                }
            });
            if (obj.modalContent.media.length === 0) {
                obj.modalContent.media.push({ type: "model", src: glbSrc });
            }
            sceneObjects.push(obj);
        });
        let envGlb = 'https://raw.githubusercontent.com/decentralize-dfw/vea-files/main/exhbt22-opt.glb';
        if (rows[0] && rows[0].c[9] && rows[0].c[9].v) {
            envGlb = rows[0].c[9].v;
        }
        let envPos = { x: 0, y: -560, z: 0 };
        let envRot = { x: 0, y: 0, z: 0 };
        let envScale = 400;
        if (rows.length > 0 && rows[0]) {
            const envRow = rows[0];
            envPos = {
                x: parseFloat(getVal(envRow, 27, 'number')),
                y: parseFloat(getVal(envRow, 28, 'number')),
                z: parseFloat(getVal(envRow, 29, 'number'))
            };
            envRot = {
                x: parseFloat(getVal(envRow, 30, 'number')),
                y: parseFloat(getVal(envRow, 31, 'number')),
                z: parseFloat(getVal(envRow, 32, 'number'))
            };
            envScale = parseFloat(getVal(envRow, 33, 'number')) || 400;
        }
        const envOptionMap = [
            { name: 'J1', row: 0, col: 9 },
            { name: 'J2', row: 1, col: 9 },
            { name: 'J3', row: 2, col: 9 },
            { name: 'J4', row: 3, col: 9 },
            { name: 'L1', row: 0, col: 11 },
            { name: 'L2', row: 1, col: 11 },
            { name: 'L3', row: 2, col: 11 },
            { name: 'L4', row: 3, col: 11 },
            { name: 'N1', row: 0, col: 13 },
            { name: 'N2', row: 1, col: 13 },
            { name: 'N3', row: 2, col: 13 },
            { name: 'N4', row: 3, col: 13 },
            { name: 'P1', row: 0, col: 15 },
            { name: 'P2', row: 1, col: 15 },
            { name: 'P3', row: 2, col: 15 },
            { name: 'P4', row: 3, col: 15 }
        ];
        const envOptions = [];
        for (let i = 0; i < envOptionMap.length; i++) {
            const m = envOptionMap[i];
            if (!rows[m.row] || !rows[m.row].c[m.col] || !rows[m.row].c[m.col].v ||
                rows[m.row].c[m.col].v === 'null' || rows[m.row].c[m.col].v.toString().trim() === '') {
                envOptions.push(null);
            } else {
                envOptions.push({ name: m.name, src: rows[m.row].c[m.col].v });
            }
        }
        sceneObjects.push({
            type: 'model',
            src: envGlb,
            position: envPos,
            rotation: envRot,
            scale: envScale,
            clickable: false,
            envOptions: envOptions,
            modalContent: { name: "Environment", artist: "VEA", description: "Environment" }
        });
    } catch (err) {
        console.error("Local Excel Load Error:", err);
        const loaderText = document.querySelector('.loader-text');
        if (loaderText) loaderText.innerText = 'DATA LOAD FAILED - CHECK FILE';
        const loaderBar = document.querySelector('.loader-bar');
        if (loaderBar) loaderBar.style.background = '#ff4444';
    }
}
const MOBILE_BREAKPOINT = 768;
const DRACO_DECODER_PATH = 'https://www.gstatic.com/draco/v1/decoders/';
function isMobileView() { return window.innerWidth <= MOBILE_BREAKPOINT; }
const getParam = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
const getParamNum = (name) => parseFloat(getParam(name));
const getParamBool = (name) => getParam(name).toLowerCase() === 'yes';
const urlParams = new URLSearchParams(window.location.search);
const RENDER_PARAMS = {
    ssaoEnabledDesktop: getParamBool('--ssao-enabled-desktop'),
    ssaoEnabledMobile: getParamBool('--ssao-enabled-mobile'),
    ssaoKernelRadius: getParamNum('--ssao-kernel-radius') || 8,
    ssaoIntensity: getParamNum('--ssao-intensity') || 0.5,
    bloomEnabledDesktop: getParamBool('--bloom-enabled-desktop'),
    bloomEnabledMobile: getParamBool('--bloom-enabled-mobile'),
    bloomStrength: getParamNum('--bloom-strength') || 0.05,
    bloomRadius: getParamNum('--bloom-radius') || 0.35,
    bloomThreshold: getParamNum('--bloom-threshold') || 0.8,
    tonemappingEnabledDesktop: getParamBool('--tonemapping-enabled-desktop'),
    tonemappingEnabledMobile: getParamBool('--tonemapping-enabled-mobile'),
    tonemappingExposureDesktop: getParamNum('--tonemapping-exposure-desktop') || 0.75,
    tonemappingExposureMobile: getParamNum('--tonemapping-exposure-mobile') || 1.0,
    shadowsEnabledDesktop: getParamBool('--shadows-enabled-desktop'),
    shadowsEnabledMobile: getParamBool('--shadows-enabled-mobile'),
    shadowType: getParam('--shadow-type') || 'PCF',
    shadowBias: getParamNum('--shadow-bias') || -0.0005,
    shadowNormalBias: getParamNum('--shadow-normal-bias') || 0.2,
    shadowMapSizeDesktop: getParamNum('--shadow-map-size-desktop') || 2048,
    shadowMapSizeMobile: getParamNum('--shadow-map-size-mobile') || 1536,
    fxaaEnabledDesktop: getParamBool('--fxaa-enabled-desktop'),
    fxaaEnabledMobile: getParamBool('--fxaa-enabled-mobile'),
    anisotropyEnabledDesktop: getParamBool('--anisotropy-enabled-desktop'),
    anisotropyEnabledMobile: getParamBool('--anisotropy-enabled-mobile'),
    envmapEnabledDesktop: getParamBool('--envmap-enabled-desktop'),
    envmapEnabledMobile: getParamBool('--envmap-enabled-mobile'),
    envmapIntensityDesktop: getParamNum('--envmap-intensity-desktop') || 1.0,
    envmapIntensityMobile: getParamNum('--envmap-intensity-mobile') || 1.0,
    rectarealightEnabledDesktop: getParamBool('--rectarealight-enabled-desktop'),
    rectarealightEnabledMobile: getParamBool('--rectarealight-enabled-mobile'),
    rectarealightIntensity: getParamNum('--rectarealight-intensity') || 2.0,
    lightHemisphereIntensity: getParamNum('--light-hemisphere-intensity') || 0.3,
    lightSunIntensity: getParamNum('--light-sun-intensity') || 3.5,
    lightSunPosX: getParamNum('--light-sun-pos-x') || 1200,
    lightSunPosY: getParamNum('--light-sun-pos-y') || 2400,
    lightSunPosZ: getParamNum('--light-sun-pos-z') || 1200,
    lightFillIntensity: getParamNum('--light-fill-intensity') || 0.1,
    lightFillPosX: getParamNum('--light-fill-pos-x') || 3000,
    lightFillPosY: getParamNum('--light-fill-pos-y') || 1500,
    lightFillPosZ: getParamNum('--light-fill-pos-z') || 3000,
    lightRimIntensity: getParamNum('--light-rim-intensity') || 0,
    lightRimPosX: getParamNum('--light-rim-pos-x') || -3000,
    lightRimPosY: getParamNum('--light-rim-pos-y') || 1500,
    lightRimPosZ: getParamNum('--light-rim-pos-z') || -3000,
    lightAccentIntensity: getParamNum('--light-accent-intensity') || 0.1,
    lightAccentPosX: getParamNum('--light-accent-pos-x') || 0,
    lightAccentPosY: getParamNum('--light-accent-pos-y') || 2000,
    lightAccentPosZ: getParamNum('--light-accent-pos-z') || -2000,
    pbrEnvmapIntensity: getParamNum('--pbr-envmap-intensity') || 1.0,
    pbrRoughnessBoost: getParamNum('--pbr-roughness-boost') || 0,
    pbrMetalnessBoost: getParamNum('--pbr-metalness-boost') || 0,
    filterBrightness: getParamNum('--filter-brightness') || 1.0,
    filterContrast: getParamNum('--filter-contrast') || 1.0,
    filterSaturation: getParamNum('--filter-saturation') || 1.0,
    filterHueRotate: getParamNum('--filter-hue-rotate') || 0,
    cam2ShowPublic: getParamBool('--cam2-show-public'),
    cam3ShowPublic: getParamBool('--cam3-show-public'),
    mobileLabelsShowPublic: getParamBool('--mobile-labels-show-public'),
    hdriPreset: getParamNum('--hdri-preset') || 1,
    hdriPresetMobile: getParamNum('--hdri-preset-mobile') || 1,
    cam1PosX: getParamNum('--cam1-pos-x'),
    cam1PosY: getParamNum('--cam1-pos-y'),
    cam1PosZ: getParamNum('--cam1-pos-z'),
    cam1LockMode: getParam('--cam1-lock-mode') || 'none',
    cam2PosX: getParamNum('--cam2-pos-x'),
    cam2PosY: getParamNum('--cam2-pos-y'),
    cam2PosZ: getParamNum('--cam2-pos-z'),
    cam2LockMode: getParam('--cam2-lock-mode') || 'none',
    cam3PosX: getParamNum('--cam3-pos-x'),
    cam3PosY: getParamNum('--cam3-pos-y'),
    cam3PosZ: getParamNum('--cam3-pos-z'),
    cam3LockMode: getParam('--cam3-lock-mode') || 'none',
    camMobilePosX: getParamNum('--cam-mobile-pos-x'),
    camMobilePosY: getParamNum('--cam-mobile-pos-y'),
    camMobilePosZ: getParamNum('--cam-mobile-pos-z'),
    camMobileLockMode: getParam('--cam-mobile-lock-mode') || 'none',
    camMobileFov: getParamNum('--cam-mobile-fov'),
    initialCameraZoom: getParamNum('--initial-camera-zoom'),
    initialCameraZoomMobile: getParamNum('--initial-camera-zoom-mobile'),
    cameraLockShowPublic: getParamBool('--camera-lock-show-public'),
    cameraDefaultLocked: getParamBool('--camera-default-locked'),
    zoomRotateEnabled: getParamBool('--zoom-rotate-enabled'),
    hoverLabelFontSize: getParamNum('--hover-label-font-size'),
    hoverLabelOffsetY: getParamNum('--hover-label-offset-y'),
    hoverLabelNameOffsetY: getParamNum('--hover-label-name-offset-y'),
    panelWidth: getParamNum('--panel-width') || 280,
    panelTopOffset: getParamNum('--panel-top-offset') || 60,
    cameraMinPolarAngle: getParamNum('--camera-min-polar-angle') || 0,
    cameraMaxPolarAngle: getParamNum('--camera-max-polar-angle') || 80,
    envOption: getParamNum('--env-option') || 1
};
const STORAGE_KEY_RENDER = 'vea_editor_renderParams';
const STORAGE_KEY_SCENE = 'vea_editor_sceneParams';
function saveParamsToStorage() {
    try {
        localStorage.setItem(STORAGE_KEY_RENDER, JSON.stringify(runtimeParams));
        localStorage.setItem(STORAGE_KEY_SCENE, JSON.stringify(PARAMS));
    } catch(e) {}
}
function loadParamsFromStorage() {
    try {
        const savedRender = localStorage.getItem(STORAGE_KEY_RENDER);
        const savedScene = localStorage.getItem(STORAGE_KEY_SCENE);
        if (savedRender) {
            const parsed = JSON.parse(savedRender);
            Object.assign(runtimeParams, parsed);
        }
        if (savedScene) {
            const parsed = JSON.parse(savedScene);
            delete parsed.showPrice;
            Object.assign(PARAMS, parsed);
        }
    } catch(e) { console.warn('Failed to load saved params:', e); }
}
function clearSavedParams() {
    try {
        localStorage.removeItem(STORAGE_KEY_RENDER);
        localStorage.removeItem(STORAGE_KEY_SCENE);
    } catch(e) {}
}
let runtimeParams = { ...RENDER_PARAMS };
window.runtimeParams = runtimeParams;
const PARAMS = {
    hoverWhileSelectedFactor: getParamNum('--hover-while-selected-factor') || 0.2,
    hoverLiftDistance: getParamNum('--hover-lift-distance'),
    hoverAnimationDuration: parseFloat(getParam('--hover-animation-duration')) * 1000,
    hoverLabelFontSize: getParamNum('--hover-label-font-size'),
    hoverLabelFontStyle: getParam('--hover-label-font-style') || 'bold',
    hoverLabelLetterSpacing: getParamNum('--hover-label-letter-spacing'),
    hoverLabelColor: getParam('--hover-label-color') || '#000000',
    hoverLabelOpacity: parseFloat(getParam('--hover-label-opacity')) || 1.0,
    hoverArtistFontSize: getParamNum('--hover-artist-font-size'),
    hoverArtistFontStyle: getParam('--hover-artist-font-style') || 'normal',
    hoverArtistLetterSpacing: getParamNum('--hover-artist-letter-spacing'),
    hoverArtistColor: getParam('--hover-artist-color') || '#333333',
    hoverArtistOpacity: parseFloat(getParam('--hover-artist-opacity')) || 0.8,
    hoverLabelFontBox: getParamNum('--hover-label-font-box'),
    hoverLabelRotationX: parseFloat(getParam('--hover-label-rotation-x')),
    hoverLabelRotationY: parseFloat(getParam('--hover-label-rotation-y')),
    hoverLabelRotationZ: parseFloat(getParam('--hover-label-rotation-z')),
    hoverLabelOffsetY: getParamNum('--hover-label-offset-y'),
    hoverLabelNameOffsetY: getParamNum('--hover-label-name-offset-y'),
    showPrice: urlParams.has('show-price') ? urlParams.get('show-price') === 'yes' : getParam('--show-price') === 'yes',
    mobileToggleScale: parseFloat(getParam('--mobile-toggle-scale')) || 1,
    selectedRotateX: getParamNum('--selected-rotate-x') || 0,
    selectedRotateY: getParamNum('--selected-rotate-y') || 0.5,
    selectedRotateZ: getParamNum('--selected-rotate-z') || 0,
    selectionZoomFactorOrtho: getParamNum('--selection-zoom-factor-ortho'),
    selectionZoomFactorPerspective: getParamNum('--selection-zoom-factor-perspective'),
    selectionZoomFactorMobile: getParamNum('--selection-zoom-factor-mobile'),
    zoomInAnimationDuration: parseFloat(getParam('--zoom-in-animation-duration')) * 1000,
    zoomOutAnimationDuration: parseFloat(getParam('--zoom-out-animation-duration')) * 1000,
    initialCameraZoom: getParamNum('--initial-camera-zoom'),
    initialCameraZoomMobile: getParamNum('--initial-camera-zoom-mobile'),
    cameraPerspectiveFov: getParamNum('--camera-perspective-fov'),
    cameraOrthoFov: getParamNum('--camera-ortho-fov'),
    cameraMinDist: getParamNum('--camera-min-distance'),
    cameraMaxDist: getParamNum('--camera-max-distance'),
    enablePan: getParam('--enable-pan') === 'yes',
    mobilePanEnabled: getParam('--mobile-pan-enabled') === 'yes',
    cam1PosX: getParamNum('--cam1-pos-x'),
    cam1PosY: getParamNum('--cam1-pos-y'),
    cam1PosZ: getParamNum('--cam1-pos-z'),
    cam2PosX: getParamNum('--cam2-pos-x'),
    cam2PosY: getParamNum('--cam2-pos-y'),
    cam2PosZ: getParamNum('--cam2-pos-z'),
    cam3PosX: getParamNum('--cam3-pos-x'),
    cam3PosY: getParamNum('--cam3-pos-y'),
    cam3PosZ: getParamNum('--cam3-pos-z'),
    camLookX: getParamNum('--cam-lookat-x'),
    camLookY: getParamNum('--cam-lookat-y'),
    camLookZ: getParamNum('--cam-lookat-z')
};
window.PARAMS = PARAMS;
loadParamsFromStorage();
// CSS --cam-mobile-fov is ALWAYS source of truth (not localStorage)
runtimeParams.camMobileFov = RENDER_PARAMS.camMobileFov;
// CSS --camera-ortho-fov is ALWAYS source of truth (not localStorage)
runtimeParams.cameraOrthoFov = RENDER_PARAMS.cameraOrthoFov;
PARAMS.cameraOrthoFov = RENDER_PARAMS.cameraOrthoFov;
window.THREE = THREE;
let camera, cameraOrtho, scene, renderer, controls, controlsOrtho, raycaster, mouse, composer, globalEnvMap;
let bloomPass, ssaoPass, fxaaPass;
let lightHemi, lightSun, lightFill, lightRim, lightAccent;
let rectLight1, rectLight2;
const hdriPresets = {
    1: { name: 'Venice Sunset', url: 'https://raw.githubusercontent.com/mrdoob/three.js/dev/examples/textures/equirectangular/venice_sunset_1k.hdr' },
    2: { name: 'White Studio', url: 'https://raw.githubusercontent.com/decentralize-dfw/co-contemporaryartifacts-glb/main/white_studio_04_1k.hdr' },
    3: { name: 'Qwantani Puresky', url: 'https://raw.githubusercontent.com/decentralize-dfw/co-contemporaryartifacts-glb/main/qwantani_mid_morning_puresky_1k.hdr' },
    4: { name: 'Story Studio', url: 'https://raw.githubusercontent.com/decentralize-dfw/co-contemporaryartifacts-glb/main/story_studio_03_1k.hdr' },
    5: { name: 'White Home Studio', url: 'https://raw.githubusercontent.com/decentralize-dfw/co-contemporaryartifacts-glb/main/white_home_studio_1k.hdr' },
    6: { name: 'Studio Kominka', url: 'https://raw.githubusercontent.com/decentralize-dfw/co-contemporaryartifacts-glb/main/studio_kominka_02_1k.hdr' }
};
let currentHdriPreset = runtimeParams.hdriPreset || 1;
let currentHdriPresetMobile = runtimeParams.hdriPresetMobile || 1;
function loadHDRI(presetNum, pmremGen, forMobile) {
    const preset = hdriPresets[presetNum];
    if (!preset) return;
    if (forMobile) {
        currentHdriPresetMobile = presetNum;
        runtimeParams.hdriPresetMobile = presetNum;
    } else {
        currentHdriPreset = presetNum;
        runtimeParams.hdriPreset = presetNum;
    }
    const isMobile = isMobileView();
    const localPmrem = pmremGen || new THREE.PMREMGenerator(renderer);
    if (!pmremGen) localPmrem.compileEquirectangularShader();
    new RGBELoader().load(preset.url, (tex) => {
        tex.mapping = THREE.EquirectangularReflectionMapping;
        if (globalEnvMap) globalEnvMap.dispose();
        globalEnvMap = localPmrem.fromEquirectangular(tex).texture;
        const useEnvMap = isMobile ? runtimeParams.envmapEnabledMobile : runtimeParams.envmapEnabledDesktop;
        if (useEnvMap && !isWireMode) {
            scene.environment = globalEnvMap;
        }
        const envIntensity = isMobile ? runtimeParams.envmapIntensityMobile : runtimeParams.envmapIntensityDesktop;
        scene.traverse(child => {
            if (child.isMesh && child.material && child.material.isMeshStandardMaterial) {
                child.material.envMapIntensity = envIntensity;
                child.material.needsUpdate = true;
            }
        });
        tex.dispose();
        if (!pmremGen) localPmrem.dispose();
        console.log(`HDRI switched to: ${presetNum} - ${preset.name}`);
    });
}
let loadedAssets = 0, totalAssets = 0;
let currentHoveredObject = null, hoverLabel = null, hoverLabelArtist = null, currentSelectedObject = null;
let originalCameraState = null, currentActiveVideos = [];
let modalAnimationFrameId;
let sceneObjects = [];
let infoTextAbout = '';
let infoTextStatus = '';
let marqueeText = '';
let currentInfoPanel = null;
 const clickableObjects = [];
let boundingBoxHelpers = [];
let boundingBoxesVisible = false;
let mobileLabelsVisible = false;
let mobileStaticLabels = [];
const clickableZones = [];
let editorModeActive = false;
let transformControls = null;
let selectedEditorObject = null;
let loadedObjectsMap = new Map();
let currentCameraMode = 'ortho';
let currentCameraIndex = 1;
let savedCameraState = null;
let currentTransformMode = 'translate';
const lockedObjects = new Set(); 
let isCameraLocked = runtimeParams.cameraDefaultLocked || false;
let undoStack = [];
let redoStack = [];
const MAX_UNDO_STEPS = 50;
let originalMaterials = new Map();
let isWireMode = false;
let wasMobileView = isMobileView();
let sceneDirty = true;
let activeTweenCount = 0;
function markDirty() { sceneDirty = true; }
async function startEngine() {
    try {
        const font = new FontFace('Telegrama Render', 'url(https://raw.githubusercontent.com/decentralize-dfw/font/main/vea.ttf)');
        await font.load();
        document.fonts.add(font);
    } catch (err) { console.log('Font load error', err); }
    document.querySelector('.loader-text').innerText = "FETCHING DATA";
    await loadDataFromSheet();
    if (marqueeText) {
        initMarquee(marqueeText);
    }
    document.querySelector('.loader-text').innerText = "LOADING ASSETS";
    init(); animate();
}
startEngine();
function init() {
    const container = document.getElementById('scene-container');
    const isMobile = isMobileView();
    renderer = new THREE.WebGLRenderer({
         antialias: true,
         alpha: true,
         powerPreference: "high-performance",
        precision: isMobile ? "mediump" : "highp",
        logarithmicDepthBuffer: true
    });
    renderer.setPixelRatio(isMobile ? Math.min(window.devicePixelRatio, 1.5) : window.devicePixelRatio);
    renderer.setSize(window.innerWidth, window.innerHeight);
    renderer.domElement.addEventListener('webglcontextlost', (e) => {
        e.preventDefault();
        console.warn('WebGL context lost');
        const loaderEl = document.getElementById('loader');
        if (loaderEl) {
            loaderEl.style.display = 'flex';
            loaderEl.style.opacity = '1';
            const loaderText = document.querySelector('.loader-text');
            if (loaderText) loaderText.innerText = 'GPU CONTEXT LOST - RECOVERING...';
        }
    });
    renderer.domElement.addEventListener('webglcontextrestored', () => {
        console.log('WebGL context restored');
        applyRuntimeParameters();
        markDirty();
        const loaderEl = document.getElementById('loader');
        if (loaderEl) {
            loaderEl.style.opacity = '0';
            setTimeout(() => { loaderEl.style.display = 'none'; }, 800);
        }
    });
    renderer.outputColorSpace = THREE.SRGBColorSpace;
    renderer.toneMapping = THREE.ACESFilmicToneMapping;
    renderer.toneMappingExposure = 0.85;
     renderer.shadowMap.enabled = true;
     renderer.shadowMap.type = THREE.PCFSoftShadowMap;
    renderer.setClearAlpha(0);
    renderer.setClearColor(0xffffff, 0);
    container.appendChild(renderer.domElement);
    const aspect = window.innerWidth / window.innerHeight;
    camera = new THREE.PerspectiveCamera(PARAMS.cameraPerspectiveFov, aspect, 10, 100000);
    camera.position.set(5000, 5000, 5000);
    cameraOrtho = new THREE.OrthographicCamera(-1, 1, 1, -1, 1, 100000);
    if (isMobile) {
        cameraOrtho.position.set(PARAMS.camMobilePosX, PARAMS.camMobilePosY, PARAMS.camMobilePosZ);
    } else {
        cameraOrtho.position.set(PARAMS.cam1PosX, PARAMS.cam1PosY, PARAMS.cam1PosZ);
    }
    cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
    if (isMobile) {
        const mobileFov = runtimeParams.camMobileFov || 45;
        cameraOrtho.zoom = PARAMS.initialCameraZoomMobile * (45 / mobileFov);
    } else {
        const orthoFov = PARAMS.cameraOrthoFov || 45;
        cameraOrtho.zoom = PARAMS.initialCameraZoom * (45 / orthoFov);
    }
    scene = new THREE.Scene();
    scene.background = null;
    setTimeout(() => {
        if (transformControls) scene.add(transformControls);
    }, 100); 
    controls = new OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.05;
    controls.enablePan = isMobile && PARAMS.mobilePanEnabled;
    controls.enableRotate = true;
    controls.minDistance = PARAMS.cameraMinDist;
    controls.maxDistance = PARAMS.cameraMaxDist;
    controls.minPolarAngle = (90 - runtimeParams.cameraMaxPolarAngle) * Math.PI / 180;
    controls.maxPolarAngle = (90 - runtimeParams.cameraMinPolarAngle) * Math.PI / 180;
    controls.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
    controls.enabled = false;
    if (isMobile && PARAMS.mobilePanEnabled) { controls.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_PAN }; }
    controlsOrtho = new OrbitControls(cameraOrtho, renderer.domElement);
    controlsOrtho.enableDamping = true;
    controlsOrtho.dampingFactor = 0.05;
    controlsOrtho.enablePan = isMobile && PARAMS.mobilePanEnabled;
    controlsOrtho.enableZoom = true;
    controlsOrtho.enableRotate = true;
    controlsOrtho.minPolarAngle = (90 - runtimeParams.cameraMaxPolarAngle) * Math.PI / 180;
    controlsOrtho.maxPolarAngle = (90 - runtimeParams.cameraMinPolarAngle) * Math.PI / 180;
    controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
    controlsOrtho.enabled = true;
    if (isMobile && PARAMS.mobilePanEnabled) { controlsOrtho.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_PAN }; }
    window.controls = controls;
    window.controlsOrtho = controlsOrtho;
    window.activeCamDot = 1;
    window.camDotFadeTimer = null;
    window.camSwitching = false;
    function onControlsChange() {
        if (isEffectivelyLocked() || editorModeActive) return;
        if (camSwitching) return;
        if (activeCamDot && !camDotFadeTimer) {
            const dot = document.getElementById('cam-dot-' + activeCamDot);
            if (dot) {
                dot.classList.add('fading');
                camDotFadeTimer = setTimeout(() => {
                    dot.classList.remove('active', 'fading');
                    activeCamDot = null;
                    camDotFadeTimer = null;
                }, 5000);
            }
        }
    }
    controls.addEventListener('change', () => { onControlsChange(); markDirty(); });
    controlsOrtho.addEventListener('change', () => { onControlsChange(); markDirty(); });
    transformControls = new TransformControls(camera, renderer.domElement);
    transformControls.setMode('translate');
    transformControls.setSize(0.8);
    transformControls.visible = false;
    transformControls.enabled = false;
    transformControls.addEventListener('dragging-changed', (event) => {
        const activeControls = currentCameraMode === 'persp' ? controls : controlsOrtho;
        activeControls.enabled = !event.value;
    });
    transformControls.addEventListener('change', () => {
        markDirty();
        if (selectedEditorObject && editorModeActive) {
            if (!selectedEditorObject.parent) {
                transformControls.detach();
                selectedEditorObject = null;
                return;
            }
            updateEditorPanelValues();
        }
    });
    transformControls.addEventListener('mouseDown', () => {
        if (selectedEditorObject && editorModeActive) {
            saveUndoState();
        }
    });
    window.addEventListener('keydown', handleKeyDown);
    const pmrem = new THREE.PMREMGenerator(renderer);
    pmrem.compileEquirectangularShader();
    const initialHdri = isMobileView() ? (runtimeParams.hdriPresetMobile || 1) : (runtimeParams.hdriPreset || 1);
    loadHDRI(initialHdri, pmrem, isMobileView());
    RectAreaLightUniformsLib.init();
    lightHemi = new THREE.HemisphereLight(0xffffff, 0x000000, runtimeParams.lightHemisphereIntensity);
    scene.add(lightHemi);
    lightSun = new THREE.DirectionalLight(0xffffff, runtimeParams.lightSunIntensity);
    lightSun.position.set(runtimeParams.lightSunPosX, runtimeParams.lightSunPosY, runtimeParams.lightSunPosZ);
    lightSun.castShadow = isMobile ? runtimeParams.shadowsEnabledMobile : runtimeParams.shadowsEnabledDesktop;
    const shadowSize = isMobile ? runtimeParams.shadowMapSizeMobile : runtimeParams.shadowMapSizeDesktop;
    lightSun.shadow.mapSize.set(shadowSize, shadowSize);
    lightSun.shadow.camera.near = 0.5;
    lightSun.shadow.camera.far = 16000;
    lightSun.shadow.camera.left = -5500; lightSun.shadow.camera.right = 5500;
    lightSun.shadow.camera.top = 5500; lightSun.shadow.camera.bottom = -5500;
    lightSun.shadow.bias = runtimeParams.shadowBias;
    lightSun.shadow.normalBias = runtimeParams.shadowNormalBias;
    scene.add(lightSun);
    lightFill = new THREE.PointLight(0xffffff, runtimeParams.lightFillIntensity, 10000);
    lightFill.position.set(runtimeParams.lightFillPosX, runtimeParams.lightFillPosY, runtimeParams.lightFillPosZ);
    scene.add(lightFill);
    lightRim = new THREE.PointLight(0xffffff, runtimeParams.lightRimIntensity, 10000);
    lightRim.position.set(runtimeParams.lightRimPosX, runtimeParams.lightRimPosY, runtimeParams.lightRimPosZ);
    scene.add(lightRim);
    lightAccent = new THREE.PointLight(0xffffff, runtimeParams.lightAccentIntensity, 8000);
    lightAccent.position.set(runtimeParams.lightAccentPosX, runtimeParams.lightAccentPosY, runtimeParams.lightAccentPosZ);
    scene.add(lightAccent);
    const useRectAreaLight = isMobile ? runtimeParams.rectarealightEnabledMobile : runtimeParams.rectarealightEnabledDesktop;
    rectLight1 = new THREE.RectAreaLight(0xffffff, runtimeParams.rectarealightIntensity, 1000, 1000);
    rectLight1.position.set(0, 3000, 3000);
    rectLight1.lookAt(0, 0, 0);
    rectLight1.visible = useRectAreaLight;
    scene.add(rectLight1);
    rectLight2 = new THREE.RectAreaLight(0xffffff, runtimeParams.rectarealightIntensity * 0.6, 800, 800);
    rectLight2.position.set(-2000, 2000, -2000);
    rectLight2.lookAt(0, 0, 0);
    rectLight2.visible = useRectAreaLight;
    scene.add(rectLight2);
    composer = new EffectComposer(renderer);
    const renderPass = new RenderPass(scene, camera);
    renderPass.clearAlpha = 0;
    composer.addPass(renderPass);
    const useSSAO = isMobile ? runtimeParams.ssaoEnabledMobile : runtimeParams.ssaoEnabledDesktop;
    ssaoPass = new SSAOPass(scene, cameraOrtho, window.innerWidth, window.innerHeight);
    ssaoPass.kernelRadius = runtimeParams.ssaoKernelRadius;
    ssaoPass.minDistance = 0.001;
    ssaoPass.maxDistance = 0.15;
    ssaoPass.output = SSAOPass.OUTPUT.Default;
    ssaoPass.enabled = useSSAO;
    composer.addPass(ssaoPass);
    const useBloom = isMobile ? runtimeParams.bloomEnabledMobile : runtimeParams.bloomEnabledDesktop;
    bloomPass = new UnrealBloomPass(new THREE.Vector2(window.innerWidth, window.innerHeight), runtimeParams.bloomStrength, runtimeParams.bloomRadius, runtimeParams.bloomThreshold);
    bloomPass.enabled = useBloom;
    composer.addPass(bloomPass);
    composer.addPass(new OutputPass());
    const useFXAA = isMobile ? runtimeParams.fxaaEnabledMobile : runtimeParams.fxaaEnabledDesktop;
    fxaaPass = new ShaderPass(FXAAShader);
    const pr = isMobile ? Math.min(window.devicePixelRatio, 1.5) : window.devicePixelRatio;
    fxaaPass.material.uniforms['resolution'].value.set(1/(window.innerWidth*pr), 1/(window.innerHeight*pr));
    fxaaPass.enabled = useFXAA;
    composer.addPass(fxaaPass);
    raycaster = new THREE.Raycaster();
    mouse = new THREE.Vector2();
    buildScene(isMobile);
    applyLightingPreset('default');
    window.addEventListener('resize', onWindowResize);
    renderer.domElement.addEventListener('mousemove', onMouseMove);
    renderer.domElement.addEventListener('mousedown', onPointerDown);
    renderer.domElement.addEventListener('mouseup', onPointerUp);
    renderer.domElement.addEventListener('click', handleEditorClick);
    renderer.domElement.addEventListener('touchstart', onTouchStart, {passive: true});
    renderer.domElement.addEventListener('touchend', onTouchEnd);
    const sidePanel = document.getElementById('side-panel');
    let touchStartY = 0;
    sidePanel.addEventListener('touchstart', (e) => {
        touchStartY = e.touches[0].clientY;
    }, {passive: true});
    sidePanel.addEventListener('touchmove', (e) => {
        const touchY = e.touches[0].clientY;
        const diffY = touchY - touchStartY;
        if (diffY > 0 && sidePanel.scrollTop <= 0) {
        }
    }, {passive: true});
    sidePanel.addEventListener('touchend', (e) => {
        const touchEndY = e.changedTouches[0].clientY;
        if (touchEndY - touchStartY > 100 && sidePanel.scrollTop <= 0) {
            closePanel();
        }
    });
    const mobileLabelsCheckbox = document.getElementById('mobile-labels-cb');
    const mobileLabelsToggle = document.getElementById('mobile-labels-toggle');
    if (mobileLabelsCheckbox) {
        mobileLabelsCheckbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                showMobileLabels();
            } else {
                hideMobileLabels();
            }
        });
    }
    if (isMobile && PARAMS.mobileLabelsShowPublic && mobileLabelsToggle) {
        mobileLabelsToggle.classList.add('visible');
    }
}
window.handlePresetSelection = function(level) { applyLightingPreset(level); };
window.applyLightingPreset = function(level) {
    markDirty();
    if (!bloomPass || !renderer) return;
    const btnLow = document.getElementById('btn-low');
    const btnMid = document.getElementById('btn-mid');
    [btnLow, btnMid].forEach(b => b?.classList.remove('active'));
    if (level === 'low') {
        isWireMode = true;
        btnLow.classList.add('active');
        bloomPass.enabled = false;
        renderer.shadowMap.enabled = false;
        renderer.toneMappingExposure = 1.0;
        renderer.toneMapping = THREE.NoToneMapping;
        scene.background = new THREE.Color(0xb6a5d0);
        scene.environment = null;
        applyLowModeMaterials(true);
        if (currentSelectedObject) {
            currentSelectedObject.traverse(child => {
                if (child.isMesh && child.material) {
                    child.material = new THREE.MeshBasicMaterial({
                        color: 0xffffff, wireframe: true,
                        transparent: true, opacity: 1, side: THREE.DoubleSide
                    });
                }
            });
        }
        document.querySelectorAll('.marquee-content').forEach(el => el.style.color = '#ffffff');
    }
    else if (level === 'default') {
        isWireMode = false;
        btnMid.classList.add('active');
        const isMobileNow = isMobileView();
        const useBloom = isMobileNow ? runtimeParams.bloomEnabledMobile : runtimeParams.bloomEnabledDesktop;
        bloomPass.enabled = useBloom;
        bloomPass.strength = runtimeParams.bloomStrength;
        const useShadows = isMobileNow ? runtimeParams.shadowsEnabledMobile : runtimeParams.shadowsEnabledDesktop;
        renderer.shadowMap.enabled = useShadows;
        renderer.shadowMap.type = THREE.PCFSoftShadowMap;
        const useTonemapping = isMobileNow ? runtimeParams.tonemappingEnabledMobile : runtimeParams.tonemappingEnabledDesktop;
        renderer.toneMapping = useTonemapping ? THREE.ACESFilmicToneMapping : THREE.NoToneMapping;
        renderer.toneMappingExposure = useTonemapping ? (isMobileNow ? runtimeParams.tonemappingExposureMobile : runtimeParams.tonemappingExposureDesktop) : 1.0;
        scene.background = null;
        const useEnvMap = isMobileNow ? runtimeParams.envmapEnabledMobile : runtimeParams.envmapEnabledDesktop;
        scene.environment = useEnvMap ? globalEnvMap : null;
        applyLowModeMaterials(false);
        const marqueeColor = getComputedStyle(document.documentElement).getPropertyValue('--marquee-color').trim() || '#b6a6d0';
        document.querySelectorAll('.marquee-content').forEach(el => el.style.color = marqueeColor);
    }
};
function applyLowModeMaterials(enableLow) {
    clickableObjects.forEach(obj => {
        obj.traverse(child => {
            if (child.isMesh && child.material) {
                if (enableLow) {
                    if (!originalMaterials.has(child.uuid)) {
                        originalMaterials.set(child.uuid, child.material);
                    }
                    const origMat = originalMaterials.get(child.uuid);
                    let color = new THREE.Color(0x29174d);
                    child.material = new THREE.MeshBasicMaterial({
                        color: color,
                        wireframe: true,
                        transparent: origMat.transparent || false,
                        opacity: origMat.opacity || 1,
                        side: origMat.side || THREE.FrontSide
                    });
                } else {
                    if (originalMaterials.has(child.uuid)) {
                        child.material = originalMaterials.get(child.uuid);
                    }
                }
            }
        });
    });
}
window.toggleWireMode = function() {
    const toggle = document.getElementById('wire-toggle');
    if (isWireMode) {
        applyLightingPreset('default');
        toggle.classList.remove('active');
        const thumb = toggle.querySelector('.wire-toggle-thumb');
        if (thumb) thumb.style.left = '0px';
    } else {
        applyLightingPreset('low');
        toggle.classList.add('active');
        const track = toggle.querySelector('.wire-toggle-track');
        const thumb = toggle.querySelector('.wire-toggle-thumb');
        if (thumb && track && track.dataset.thumbActiveLeft) {
            thumb.style.left = track.dataset.thumbActiveLeft;
        }
    }
};
function positionWireToggle() {
    positionRightToggles();
}
function positionRightToggles() {
    const rightToggles = document.getElementById('right-toggles');
    if (!rightToggles) return;
    const isMobileNow = isMobileView();
    const row0 = document.getElementById('right-toggle-row0');
    const row1 = document.getElementById('right-toggle-row1');
    const row2 = document.getElementById('right-toggle-row2');
    if (isMobileNow) {
        if (row0) { row0.style.position = 'relative'; row0.style.top = ''; }
        if (row1) { row1.style.position = 'relative'; row1.style.top = ''; }
        if (row2) { row2.style.position = 'relative'; row2.style.top = ''; }
        // Mobilde: yatay sıra, sağ hizalı, info-buttons'ın altında
        rightToggles.style.flexDirection = 'row';
        rightToggles.style.gap = '8px';
        rightToggles.style.alignItems = 'flex-start';
        rightToggles.style.right = '12px';
        rightToggles.style.left = '';
        rightToggles.style.transform = '';
        // info-buttons'ın altına konumlandır
        const infoButtons = document.getElementById('info-buttons');
        if (infoButtons) {
            const infoRect = infoButtons.getBoundingClientRect();
            rightToggles.style.top = (infoRect.bottom + 6) + 'px';
        } else {
            rightToggles.style.top = '44px';
        }
        return;
    }
    // Desktop: row0=grid/3D, row1=wire, row2=camera dots
    rightToggles.style.left = '';
    rightToggles.style.transform = '';
    rightToggles.style.alignItems = 'flex-end';
    rightToggles.style.top = '0px';
    rightToggles.style.gap = '0';
    const btnStatus = document.getElementById('btn-status');
    const btnAbout = document.getElementById('btn-about');
    // Row0 (grid/3D toggle): align vertical center with btn-status
    if (row0 && btnStatus) {
        const statusRect = btnStatus.getBoundingClientRect();
        const statusCenterY = statusRect.top + statusRect.height / 2;
        const row0Height = row0.getBoundingClientRect().height;
        row0.style.position = 'absolute';
        row0.style.top = (statusCenterY - row0Height / 2) + 'px';
    }
    // Row1 (wire toggle): align vertical center with btn-about
    if (row1 && btnAbout) {
        const aboutRect = btnAbout.getBoundingClientRect();
        const aboutCenterY = aboutRect.top + aboutRect.height / 2;
        const row1Height = row1.getBoundingClientRect().height;
        row1.style.position = 'absolute';
        row1.style.top = (aboutCenterY - row1Height / 2) + 'px';
    }
    // Row2 (camera dots): equal center-to-center spacing
    // Distance from row0-center to row1-center = distance from row1-center to row2-center
    if (row2) {
        const row0Rect = row0 ? row0.getBoundingClientRect() : null;
        const row1Rect = row1 ? row1.getBoundingClientRect() : null;
        if (row0Rect && row1Rect) {
            const row0CenterY = row0Rect.top + row0Rect.height / 2;
            const row1CenterY = row1Rect.top + row1Rect.height / 2;
            const centerSpacing = row1CenterY - row0CenterY;
            const row2Height = row2.getBoundingClientRect().height;
            row2.style.position = 'absolute';
            row2.style.top = (row1CenterY + centerSpacing - row2Height / 2) + 'px';
        } else {
            const row1Top = row1 ? parseFloat(row1.style.top) || 0 : 0;
            row2.style.position = 'absolute';
            row2.style.top = (row1Top + 30) + 'px';
        }
    }
}
function buildScene(isMobile) {
    startLoadTimeout();
    const loader = new GLTFLoader();
    const draco = new DRACOLoader();
    draco.setDecoderPath(DRACO_DECODER_PATH);
    loader.setDRACOLoader(draco);
    loader.setMeshoptDecoder(MeshoptDecoder);
    const envObj = sceneObjects.find(o => o.modalContent && o.modalContent.name === 'Environment');
    if (envObj && envObj.envOptions) {
        const optIdx = (runtimeParams.envOption || 1) - 1;
        if (envObj.envOptions[optIdx] && envObj.envOptions[optIdx].src) {
            envObj.src = envObj.envOptions[optIdx].src;
        }
    }
    const validObjects = sceneObjects.filter(item => !(isMobile && item.showOnMobile === false));
    totalAssets = validObjects.length;
    if(totalAssets === 0) {
         document.getElementById('loader').style.display = 'none';
         return;
    }
    sceneObjects.forEach((item, index) => {
        if (isMobile && item.showOnMobile === false) return;
        loader.load(item.src, (gltf) => {
            const m = gltf.scene;
            m.position.set(item.position.x, item.position.y || 0, item.position.z);
            if (item.rotation) m.rotation.set(item.rotation.x*Math.PI/180, item.rotation.y*Math.PI/180, item.rotation.z*Math.PI/180);
            m.scale.setScalar(item.scale);
            m.userData = {
                modalContent: item.modalContent,
                clickable: item.clickable !== false,
                includeInFit: item.includeInFit !== false,
                originalPosition: m.position.clone(),
                sceneObjectIndex: index,
                originalScale: item.scale,
                rowIndex: item.rowIndex,
                gridThumbnail: item.gridThumbnail || '',
                thumbScale: item.thumbScale,
                gridScale: item.gridScale,
                gridRotX: item.gridRotX,
                gridRotY: item.gridRotY,
                gridOffsetY: item.gridOffsetY
            };
            m.traverse(c => {
                if(c.isMesh) {
                    c.castShadow = true;
                    c.receiveShadow = true;
                    if(c.geometry.computeBoundsTree) c.geometry.computeBoundsTree();
                }
            });
            scene.add(m);
            clickableObjects.push(m);
            loadedObjectsMap.set(index, m);
            loadedAssets++; updateLoader();
            if(loadedAssets === totalAssets) {
                const isMobileInit = isMobileView();
                switchCamera(isMobileInit ? 'mobile' : 1);
                buildEditorPanel();
                buildBoundingBoxes();
                buildClickableZones();
                updateCam3Visibility();
            }
        }, undefined, (err) => {
            console.error("Error loading model:", item.src, err);
            loadedAssets++; updateLoader();
         });
    });
}
let loadTimeoutId = null;
function startLoadTimeout() {
    loadTimeoutId = setTimeout(() => {
        if (loadedAssets < totalAssets) {
            const loaderText = document.querySelector('.loader-text');
            if (loaderText) loaderText.innerText = "LOADING TIMEOUT - PLEASE REFRESH";
            console.warn(`Asset loading timed out: ${loadedAssets}/${totalAssets} loaded`);
        }
    }, 45000);
}
function updateLoader() {
    const p = Math.round((loadedAssets/totalAssets)*100);
    document.querySelector('.loader-bar').style.width = p+'%';
    if(loadedAssets>=totalAssets) {
        if (loadTimeoutId) { clearTimeout(loadTimeoutId); loadTimeoutId = null; }
        setTimeout(()=> {
            document.getElementById('loader').style.opacity='0';
            setTimeout(()=>{ document.getElementById('loader').style.display='none'; positionRightToggles(); applyInitialLockState(); }, 800);
        }, 500);
    }
}
function onMouseMove(e) {
    if (isMobileView()) return;
    if (editorModeActive) return;
    if (gridModeActive || gridTransitioning) return;
    mouse.x = (e.clientX/window.innerWidth)*2-1;
    mouse.y = -(e.clientY/window.innerHeight)*2+1;
    raycaster.setFromCamera(mouse, getActiveCamera());
    const zoneHits = raycaster.intersectObjects(clickableZones, false);
    if (zoneHits.length > 0) {
        const zone = zoneHits[0].object;
        const parentObj = zone.userData.parentObject;
        if (parentObj && parentObj.userData.modalContent && parentObj.userData.clickable && currentHoveredObject !== parentObj && parentObj !== currentSelectedObject) {
            clearHover();
            applyHover(parentObj);
        }
        return;
    }
    const hits = raycaster.intersectObjects(clickableObjects, true);
    if(hits.length>0) {
        let obj = hits[0].object;
        while(obj.parent && !obj.userData.modalContent) obj = obj.parent;
        if(obj.userData.modalContent && obj.userData.clickable && currentHoveredObject !== obj && obj !== currentSelectedObject) {
            clearHover();
            applyHover(obj);
        } else if (!obj.userData.clickable) {
            clearHover();
        }
    } else clearHover();
}
function wrapText(context, text, maxWidth) {
    const words = text.split(' ');
    let line = '';
    const lines = [];
    for(let n = 0; n < words.length; n++) {
        const testLine = line + words[n] + ' ';
        const metrics = context.measureText(testLine);
        const testWidth = metrics.width;
        if (testWidth > maxWidth && n > 0) {
            lines.push(line);
            line = words[n] + ' ';
        } else {
            line = testLine;
        }
    }
    lines.push(line);
    return lines;
}
function createHoverTextMesh(text, fontSize, maxW, opts = {}) {
    const fontStyle = opts.fontStyle || 'bold';
    const letterSpacing = opts.letterSpacing || 0;
    const fillColor = isWireMode ? '#ffffff' : (opts.color || '#000000');
    const opacity = opts.opacity !== undefined ? opts.opacity : 1.0;
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const fontStr = `${fontStyle} ${fontSize}px "Space Mono"`;
    ctx.font = fontStr;
    const lines = wrapText(ctx, text, maxW);
    const lineHeight = fontSize * 1.1;
    canvas.width = maxW + 100;
    canvas.height = (lines.length * lineHeight) + 80;
    ctx.font = fontStr;
    ctx.fillStyle = fillColor;
    ctx.globalAlpha = opacity;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    if (letterSpacing) ctx.letterSpacing = `${letterSpacing}px`;
    const centerX = canvas.width / 2;
    let startY = (canvas.height - (lines.length * lineHeight)) / 2 + (lineHeight / 2);
    lines.forEach((line, i) => {
        ctx.fillText(line.trim(), centerX, startY + (i * lineHeight));
    });
    const tex = new THREE.CanvasTexture(canvas);
    tex.needsUpdate = true;
    tex.anisotropy = renderer.capabilities.getMaxAnisotropy();
    const mat = new THREE.MeshBasicMaterial({
        map: tex, transparent: true,
        depthTest: false, depthWrite: false,
        side: THREE.DoubleSide, toneMapped: false
    });
    const mesh = new THREE.Mesh(new THREE.PlaneGeometry(canvas.width, canvas.height), mat);
    mesh.renderOrder = 9999;
    return mesh;
}
function applyHover(obj) {
    currentHoveredObject = obj;
    const hasSelection = !!currentSelectedObject;
    const f = hasSelection ? PARAMS.hoverWhileSelectedFactor : 1.0;
    const liftDist = PARAMS.hoverLiftDistance * f;
    new TWEEN(obj.position).to({ y: obj.userData.originalPosition.y + liftDist }, PARAMS.hoverAnimationDuration).start();
    const box = new THREE.Box3().setFromObject(obj);
    const bboxTopLocal = box.max.y - obj.position.y;
    obj.userData._bboxTopLocal = bboxTopLocal;
    obj.userData._hoverLiftMult = f;
    const maxW = PARAMS.hoverLabelFontBox;
    const content = obj.userData.modalContent;
    const rotX = PARAMS.hoverLabelRotationX * Math.PI / 180;
    const rotY = PARAMS.hoverLabelRotationY * Math.PI / 180;
    const rotZ = PARAMS.hoverLabelRotationZ * Math.PI / 180;
    const artistText = content.artist || '';
    if (artistText) {
        hoverLabelArtist = createHoverTextMesh(artistText, PARAMS.hoverArtistFontSize, maxW, {
            fontStyle: PARAMS.hoverArtistFontStyle,
            letterSpacing: PARAMS.hoverArtistLetterSpacing,
            color: PARAMS.hoverArtistColor,
            opacity: PARAMS.hoverArtistOpacity
        });
        const artistY = obj.position.y + bboxTopLocal + liftDist + PARAMS.hoverLabelOffsetY * f;
        hoverLabelArtist.position.set(obj.position.x, artistY, obj.position.z);
        hoverLabelArtist.rotation.set(rotX, rotY, rotZ);
        hoverLabelArtist.scale.setScalar(f);
        scene.add(hoverLabelArtist);
    }
    const nameText = content.hoverDisplayName || content.name || '';
    if (nameText) {
        hoverLabel = createHoverTextMesh(nameText, PARAMS.hoverLabelFontSize, maxW, {
            fontStyle: PARAMS.hoverLabelFontStyle,
            letterSpacing: PARAMS.hoverLabelLetterSpacing,
            color: PARAMS.hoverLabelColor,
            opacity: PARAMS.hoverLabelOpacity
        });
        const nameY = obj.position.y + bboxTopLocal + liftDist + (PARAMS.hoverLabelOffsetY + PARAMS.hoverLabelNameOffsetY) * f;
        hoverLabel.position.set(obj.position.x, nameY, obj.position.z);
        hoverLabel.rotation.set(rotX, rotY, rotZ);
        hoverLabel.scale.setScalar(f);
        scene.add(hoverLabel);
    }
}
function clearHover() {
    if(currentHoveredObject) {
        new TWEEN(currentHoveredObject.position).to({ y: currentHoveredObject.userData.originalPosition.y }, PARAMS.hoverAnimationDuration).start();
        currentHoveredObject = null;
    }
    if(hoverLabel) { scene.remove(hoverLabel); if(hoverLabel.material) { if(hoverLabel.material.map) hoverLabel.material.map.dispose(); hoverLabel.material.dispose(); } if(hoverLabel.geometry) hoverLabel.geometry.dispose(); hoverLabel=null; }
    if(hoverLabelArtist) { scene.remove(hoverLabelArtist); if(hoverLabelArtist.material) { if(hoverLabelArtist.material.map) hoverLabelArtist.material.map.dispose(); hoverLabelArtist.material.dispose(); } if(hoverLabelArtist.geometry) hoverLabelArtist.geometry.dispose(); hoverLabelArtist=null; }
}
let sX, sY;
function onPointerDown(e) { sX=e.clientX; sY=e.clientY; }
function onPointerUp(e) {
    if (editorModeActive) return;
    if(Math.abs(e.clientX-sX)<5 && Math.abs(e.clientY-sY)<5) {
        mouse.x = (e.clientX/window.innerWidth)*2-1;
        mouse.y = -(e.clientY/window.innerHeight)*2+1;
        raycaster.setFromCamera(mouse, getActiveCamera());
        const zoneHits = raycaster.intersectObjects(clickableZones, false);
        if (zoneHits.length > 0) {
            const zone = zoneHits[0].object;
            const parentObj = zone.userData.parentObject;
            if (parentObj && parentObj.userData.modalContent && parentObj.userData.clickable) {
                openPanel(parentObj.userData.modalContent, parentObj);
                return;
            }
        }
        const hits = raycaster.intersectObjects(clickableObjects, true);
        if(hits.length>0) {
            let obj = hits[0].object;
            while(obj.parent && !obj.userData.modalContent) obj = obj.parent;
            if(obj.userData.modalContent && obj.userData.clickable) {
                openPanel(obj.userData.modalContent, obj);
            }
        }
    }
}
let touchSX, touchSY;
function onTouchStart(e) {
    if (e.touches.length !== 1) return;
    touchSX = e.touches[0].clientX;
    touchSY = e.touches[0].clientY;
}
function onTouchEnd(e) {
    if (editorModeActive) return;
    if (gridModeActive || gridTransitioning) return;
    if (document.getElementById('side-panel').classList.contains('open')) return;
    const t = e.changedTouches[0];
    if (Math.abs(t.clientX - touchSX) > 10 || Math.abs(t.clientY - touchSY) > 10) return;
    mouse.x = (t.clientX / window.innerWidth) * 2 - 1;
    mouse.y = -(t.clientY / window.innerHeight) * 2 + 1;
    raycaster.setFromCamera(mouse, getActiveCamera());
    // Mobile: ONLY use bounding box zones for click detection (no mesh fallback)
    const zoneHits = raycaster.intersectObjects(clickableZones, false);
    if (zoneHits.length > 0) {
        const zone = zoneHits[0].object;
        const parentObj = zone.userData.parentObject;
        if (parentObj && parentObj.userData.modalContent && parentObj.userData.clickable) {
            openPanel(parentObj.userData.modalContent, parentObj);
        }
    }
}
function formatEuroPrice(val) {
    return String(val);
}
function openPanel(content, obj) {
    if(currentSelectedObject === obj && document.getElementById('side-panel').classList.contains('open')) return;
    if(hoverLabel) { scene.remove(hoverLabel); if(hoverLabel.material) { if(hoverLabel.material.map) hoverLabel.material.map.dispose(); hoverLabel.material.dispose(); } if(hoverLabel.geometry) hoverLabel.geometry.dispose(); hoverLabel=null; }
    if(hoverLabelArtist) { scene.remove(hoverLabelArtist); if(hoverLabelArtist.material) { if(hoverLabelArtist.material.map) hoverLabelArtist.material.map.dispose(); hoverLabelArtist.material.dispose(); } if(hoverLabelArtist.geometry) hoverLabelArtist.geometry.dispose(); hoverLabelArtist=null; }
    currentHoveredObject = null;
    currentActiveVideos.forEach(v => { v.pause(); v.src = ""; v.load(); });
    currentActiveVideos = [];
    const activeCam = getActiveCamera();
    const activeControls = currentCameraMode === 'persp' ? controls : controlsOrtho;
    if(!originalCameraState) {
        originalCameraState = {
            zoom: activeCam.zoom,
            target: activeControls.target.clone(),
            position: activeCam.position.clone(),
            locked: isEffectivelyLocked()
        };
    }
    if (currentSelectedObject && currentSelectedObject !== obj) {
        if (currentSelectedObject.userData._originalRotation) {
            const origRot = currentSelectedObject.userData._originalRotation;
            currentSelectedObject.rotation.set(origRot.x, origRot.y, origRot.z);
            currentSelectedObject.userData._originalRotation = null;
        }
        if (currentSelectedObject.userData.originalPosition) {
            currentSelectedObject.position.y = currentSelectedObject.userData.originalPosition.y;
        }
        if (isWireMode) {
            currentSelectedObject.traverse(child => {
                if (child.isMesh && originalMaterials.has(child.uuid)) {
                    const origMat = originalMaterials.get(child.uuid);
                    child.material = new THREE.MeshBasicMaterial({
                        color: 0x29174d, wireframe: true,
                        transparent: origMat.transparent || false,
                        opacity: origMat.opacity || 1,
                        side: origMat.side || THREE.FrontSide
                    });
                }
            });
        }
    }
    currentSelectedObject = obj;
    if (!obj.userData._originalRotation) {
        obj.userData._originalRotation = obj.rotation.clone();
    }
    if (isWireMode) {
        obj.traverse(child => {
            if (child.isMesh && child.material) {
                child.material = new THREE.MeshBasicMaterial({
                    color: 0xffffff, wireframe: true,
                    transparent: true, opacity: 1, side: THREE.DoubleSide
                });
            }
        });
    }
    const box = new THREE.Box3().setFromObject(obj);
    const center = box.getCenter(new THREE.Vector3());
    const isMobileNow = isMobileView();
    const mFovSel = runtimeParams.camMobileFov || 45;
    const dOrthoFov = runtimeParams.cameraOrthoFov || 45;
    const baseZoom = isMobileNow ? (PARAMS.initialCameraZoomMobile * (45 / mFovSel)) : (PARAMS.initialCameraZoom * (45 / dOrthoFov));
    const zoomFactor = isMobileNow ? PARAMS.selectionZoomFactorMobile : (currentCameraMode === 'persp' ? PARAMS.selectionZoomFactorPerspective : PARAMS.selectionZoomFactorOrtho);
    const targetZoom = baseZoom / zoomFactor;
    if (isMobileNow) {
        // Mobile: instant zoom, no animation (TWEEN with dur=0 causes NaN)
        activeCam.zoom = targetZoom;
        activeCam.updateProjectionMatrix();
        activeControls.target.set(center.x, center.y, center.z);
        activeControls.update();
        markDirty();
    } else {
        new TWEEN(activeCam)
            .to({ zoom: targetZoom }, PARAMS.zoomInAnimationDuration)
            .onUpdate(() => activeCam.updateProjectionMatrix())
            .start();
        new TWEEN(activeControls.target)
            .to({ x: center.x, y: center.y, z: center.z }, PARAMS.zoomInAnimationDuration)
            .start();
    }
    applyLockState();
    document.getElementById('panel-artist').innerText = content.artist || "UNKNOWN";
    document.getElementById('panel-title').innerText = content.name;
    const metaGrid = document.getElementById('panel-meta-grid');
    const val = (v) => (!v || v === "") ? '-' : v;
    let metaHTML = `
        <span class="meta-label">DIMENSIONS:</span> <span class="meta-value">${val(content.dimensions)}</span>
        <span class="meta-label">CATEGORY:</span> <span class="meta-value">${val(content.category)}</span>
        <span class="meta-label">PRODUCT TYPE:</span> <span class="meta-value">${val(content.type)}</span>
        <span class="meta-label">COLOR:</span> <span class="meta-value">${val(content.color)}</span>
    `;
    if (content.edition && content.edition !== "") {
        metaHTML += `<span class="meta-label">EDITION:</span> <span class="meta-value">${content.edition}</span>`;
    }
    if (PARAMS.showPrice && content.price && content.price !== "") {
        const formattedPrice = formatEuroPrice(content.price);
        metaHTML += `<span class="meta-label">PRICE:</span> <span class="meta-value">${formattedPrice}</span>`;
    }
    metaGrid.innerHTML = metaHTML;
    document.getElementById('panel-desc').innerText = content.description || "";
    document.getElementById('panel-bio-header').innerText = content.artist ? `ABOUT ${content.artist}` : "";
    document.getElementById('panel-bio').innerText = content.bio || "";
    const linksContainer = document.getElementById('panel-links');
     linksContainer.innerHTML = '';
     const mediaContainer = document.getElementById('panel-media');
     mediaContainer.innerHTML = '';
    const images = content.media.filter(m => m.type === 'image');
    const models = content.media.filter(m => m.type === 'model');
    if (images.length > 0) {
        let btnLeft = null;
        let btnRight = null;
        const carouselDiv = document.createElement('div');
        carouselDiv.className = 'media-carousel-container';
        let currentImgIdx = 0;
        images.forEach((imgData, idx) => {
            const img = document.createElement('img');
            img.src = imgData.src;
            img.className = `carousel-img ${idx === 0 ? 'active' : ''}`;
            img.dataset.index = idx;
            carouselDiv.appendChild(img);
        });
        let dotsDiv = null;
        const updateCarousel = () => {
            Array.from(carouselDiv.querySelectorAll('.carousel-img')).forEach(img => img.classList.remove('active'));
            const imgElements = carouselDiv.querySelectorAll('.carousel-img');
            if (imgElements[currentImgIdx]) imgElements[currentImgIdx].classList.add('active');
            if(images.length > 1) {
                if(btnLeft) btnLeft.style.display = currentImgIdx > 0 ? 'flex' : 'none';
                if(btnRight) btnRight.style.display = currentImgIdx < images.length - 1 ? 'flex' : 'none';
                if (dotsDiv) {
                    dotsDiv.querySelectorAll('.carousel-dot').forEach((dot, i) => {
                        dot.classList.toggle('active', i === currentImgIdx);
                    });
                }
            }
        };
        if (images.length > 1) {
            btnLeft = document.createElement('div');
            btnLeft.className = 'carousel-nav-btn nav-left';
            btnLeft.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>`;
            btnLeft.onclick = (e) => {
                e.stopPropagation();
                if (currentImgIdx > 0) { currentImgIdx--; updateCarousel(); }
            };
            btnRight = document.createElement('div');
            btnRight.className = 'carousel-nav-btn nav-right';
            btnRight.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>`;
            btnRight.onclick = (e) => {
                e.stopPropagation();
                if (currentImgIdx < images.length - 1) { currentImgIdx++; updateCarousel(); }
            };
            dotsDiv = document.createElement('div');
            dotsDiv.className = 'carousel-dots';
            images.forEach((_, idx) => {
                const dot = document.createElement('div');
                dot.className = `carousel-dot ${idx === 0 ? 'active' : ''}`;
                dot.onclick = (e) => {
                    e.stopPropagation();
                    currentImgIdx = idx;
                    updateCarousel();
                };
                dotsDiv.appendChild(dot);
            });
            updateCarousel();
        }
        let touchStartX = 0;
        carouselDiv.addEventListener('touchstart', (e) => {
            touchStartX = e.touches[0].clientX;
        }, {passive: true});
        carouselDiv.addEventListener('touchend', (e) => {
            const touchEndX = e.changedTouches[0].clientX;
            const diff = touchEndX - touchStartX;
            if (Math.abs(diff) > 50) {
                if (diff > 0 && currentImgIdx > 0) {
                    currentImgIdx--;
                    updateCarousel();
                } else if (diff < 0 && currentImgIdx < images.length - 1) {
                    currentImgIdx++;
                    updateCarousel();
                }
            }
        });
        mediaContainer.appendChild(carouselDiv);
        if (btnLeft) mediaContainer.appendChild(btnLeft);
        if (btnRight) mediaContainer.appendChild(btnRight);
        if (dotsDiv) mediaContainer.appendChild(dotsDiv);
        const positionCarouselArrows = () => {
            if (!carouselDiv.offsetHeight) return;
            const centerY = carouselDiv.offsetTop + carouselDiv.offsetHeight / 2;
            if (btnLeft) btnLeft.style.top = centerY + 'px';
            if (btnRight) btnRight.style.top = centerY + 'px';
        };
        requestAnimationFrame(() => requestAnimationFrame(positionCarouselArrows));
        if (window._carouselResizeHandler) window.removeEventListener('resize', window._carouselResizeHandler);
        window._carouselResizeHandler = positionCarouselArrows;
        window.addEventListener('resize', positionCarouselArrows);
    }
    models.forEach(m => {
        const item = document.createElement('div');
         item.className = 'panel-media-item';
        let wrapper = document.createElement('div');
        wrapper.className = 'media-wrapper';
        wrapper.style.height = '300px';
        item.appendChild(wrapper);
        mediaContainer.appendChild(item);
        setTimeout(() => initModal3D(wrapper, m.src), 50);
    });
    document.getElementById('side-panel').classList.add('open');
    document.getElementById('side-panel').scrollTop = 0;
}
function initModal3D(container, src) {
    const sceneM = new THREE.Scene();
     sceneM.background = new THREE.Color(0xf4f4f4);
     if (globalEnvMap) sceneM.environment = globalEnvMap;
    const renM = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    renM.setPixelRatio(window.devicePixelRatio);
    renM.toneMapping = THREE.ACESFilmicToneMapping;
    renM.toneMappingExposure = 1.3;
    container.appendChild(renM.domElement);
    const camM = new THREE.PerspectiveCamera(45, 1, 0.1, 10000);
    sceneM.add(new THREE.HemisphereLight(0xffffff, 0xffffff, 1.2));
    const dirLight = new THREE.DirectionalLight(0xffffff, 3.5);
     dirLight.position.set(5, 10, 7);
     sceneM.add(dirLight);
    const ctrlM = new OrbitControls(camM, renM.domElement);
     ctrlM.autoRotate = true; ctrlM.autoRotateSpeed = 2.5;
     ctrlM.enableDamping = true;
    const loader = new GLTFLoader();
    const draco = new DRACOLoader();
    draco.setDecoderPath(DRACO_DECODER_PATH);
    loader.setDRACOLoader(draco);
    loader.setMeshoptDecoder(MeshoptDecoder);
    loader.load(src, (gltf) => {
        const model = gltf.scene;
        model.traverse(c => { if (c.isMesh) { c.castShadow = true; c.receiveShadow = true; } });
        sceneM.add(model);
        const box = new THREE.Box3().setFromObject(model);
        const center = box.getCenter(new THREE.Vector3());
        const size = box.getSize(new THREE.Vector3());
        const maxDim = Math.max(size.x, size.y, size.z);
        ctrlM.target.copy(center);
        camM.position.set(center.x, center.y + maxDim * 0.2, center.z + maxDim * 2.5);
        const w = container.offsetWidth || 300;
        renM.setSize(w, w);
        camM.aspect = 1;
        camM.updateProjectionMatrix();
        ctrlM.update();
    });
    function disposeModal3D() {
        ctrlM.dispose();
        sceneM.traverse(c => {
            if (c.geometry) c.geometry.dispose();
            if (c.material) {
                if (Array.isArray(c.material)) c.material.forEach(m => m.dispose());
                else c.material.dispose();
            }
        });
        renM.dispose();
        if (renM.domElement && renM.domElement.parentNode) {
            renM.domElement.parentNode.removeChild(renM.domElement);
        }
    }
    function loop() {
         if (!document.getElementById('side-panel').classList.contains('open')) { disposeModal3D(); return; }
        modalAnimationFrameId = requestAnimationFrame(loop);
         ctrlM.update();
         renM.render(sceneM, camM);
     }
    loop();
}
window.closePanel = function() {
    currentActiveVideos.forEach(v => { v.pause(); v.src = ""; v.load(); v.remove(); });
    currentActiveVideos = [];
    document.getElementById('panel-media').innerHTML = "";
    const isMobileNow = isMobileView();
    if(originalCameraState) {
        const activeCam = getActiveCamera();
        const activeControls = currentCameraMode === 'persp' ? controls : controlsOrtho;
        if (isMobileNow) {
            // Mobile: instant restore (no animation behind panel)
            activeCam.zoom = originalCameraState.zoom;
            activeCam.updateProjectionMatrix();
            activeControls.target.copy(originalCameraState.target);
            if (originalCameraState.position) {
                activeCam.position.copy(originalCameraState.position);
            }
            activeControls.update();
            applyLockState();
        } else {
            new TWEEN(activeCam)
                .to({ zoom: originalCameraState.zoom }, PARAMS.zoomOutAnimationDuration)
                .onUpdate(() => activeCam.updateProjectionMatrix())
                .onComplete(() => { applyLockState(); })
                .start();
            new TWEEN(activeControls.target)
                .to({ x: originalCameraState.target.x, y: originalCameraState.target.y, z: originalCameraState.target.z }, PARAMS.zoomOutAnimationDuration)
                .start();
            if (isEffectivelyLocked() && originalCameraState.position) {
                new TWEEN(activeCam.position)
                    .to({ x: originalCameraState.position.x, y: originalCameraState.position.y, z: originalCameraState.position.z }, PARAMS.zoomOutAnimationDuration)
                    .start();
            }
        }
    }
    if(currentSelectedObject) {
        if (currentSelectedObject.userData.originalPosition) {
            if (isMobileNow) {
                currentSelectedObject.position.y = currentSelectedObject.userData.originalPosition.y;
            } else {
                new TWEEN(currentSelectedObject.position).to({ y: currentSelectedObject.userData.originalPosition.y }, PARAMS.hoverAnimationDuration).start();
            }
        }
        if (currentSelectedObject.userData._originalRotation) {
            const origRot = currentSelectedObject.userData._originalRotation;
            if (isMobileNow) {
                currentSelectedObject.rotation.set(origRot.x, origRot.y, origRot.z);
            } else {
                new TWEEN(currentSelectedObject.rotation).to({ x: origRot.x, y: origRot.y, z: origRot.z }, PARAMS.hoverAnimationDuration).start();
            }
            currentSelectedObject.userData._originalRotation = null;
        }
        if (isWireMode) {
            currentSelectedObject.traverse(child => {
                if (child.isMesh && originalMaterials.has(child.uuid)) {
                    child.material = new THREE.MeshBasicMaterial({
                        color: 0x29174d, wireframe: true,
                        transparent: originalMaterials.get(child.uuid).transparent || false,
                        opacity: originalMaterials.get(child.uuid).opacity || 1,
                        side: originalMaterials.get(child.uuid).side || THREE.FrontSide
                    });
                }
            });
        }
    }
    originalCameraState = null; currentSelectedObject=null;
    markDirty();
    clearHover();
    if(modalAnimationFrameId) { cancelAnimationFrame(modalAnimationFrameId); modalAnimationFrameId = null; }
    document.getElementById('side-panel').classList.remove('open');
}

// ==================== GRID MODAL (Contemporary Artifacts Style) ====================
let gmCarouselImages = [];
let gmCarouselIndex = 0;

function gmUpdateCarousel() {
    const img = document.getElementById('gm-main-img');
    const thumbs = document.getElementById('gm-thumbnails');
    const dots = document.getElementById('gm-dots');
    const prev = document.getElementById('gm-prev');
    const next = document.getElementById('gm-next');
    const hasMulti = gmCarouselImages.length > 1;
    const isFirst = gmCarouselIndex === 0;
    const isLast = gmCarouselIndex === gmCarouselImages.length - 1;

    if (img) img.src = gmCarouselImages[gmCarouselIndex];
    if (prev) prev.classList.toggle('hidden', !hasMulti || isFirst);
    if (next) next.classList.toggle('hidden', !hasMulti || isLast);

    if (hasMulti && thumbs) {
        thumbs.classList.add('show');
        thumbs.innerHTML = gmCarouselImages.map((src, idx) =>
            `<div class="gm-thumb ${idx === gmCarouselIndex ? 'active' : ''}" onclick="window.gmCarouselGoTo(${idx})"><img src="${src}" alt=""></div>`
        ).join('');
    } else if (thumbs) {
        thumbs.classList.remove('show');
        thumbs.innerHTML = '';
    }

    if (hasMulti && dots) {
        dots.classList.add('show');
        dots.innerHTML = gmCarouselImages.map((_, idx) =>
            `<div class="gm-dot ${idx === gmCarouselIndex ? 'active' : ''}" onclick="window.gmCarouselGoTo(${idx})"></div>`
        ).join('');
    } else if (dots) {
        dots.classList.remove('show');
        dots.innerHTML = '';
    }
}

window.gmCarouselNext = function() {
    if (gmCarouselIndex < gmCarouselImages.length - 1) { gmCarouselIndex++; gmUpdateCarousel(); }
};
window.gmCarouselPrev = function() {
    if (gmCarouselIndex > 0) { gmCarouselIndex--; gmUpdateCarousel(); }
};
window.gmCarouselGoTo = function(idx) {
    if (idx >= 0 && idx < gmCarouselImages.length) { gmCarouselIndex = idx; gmUpdateCarousel(); }
};

window.openGridModal = function(content) {
    // Setup carousel images
    gmCarouselImages = (content.media || []).filter(m => m.type === 'image').map(m => m.src);
    if (gmCarouselImages.length === 0) gmCarouselImages = [''];
    gmCarouselIndex = 0;
    gmUpdateCarousel();

    // Build detail content
    const val = (v) => (!v || v === '') ? null : v;
    let html = '';
    html += '<div style="display:flex;flex-direction:column;align-items:flex-start;gap:12px;margin-bottom:2.5rem;">';
    html += `<h2 style="font-family:\'Capriola\',sans-serif;font-size:18pt;font-weight:bold;text-transform:uppercase;line-height:1.1;letter-spacing:-0.5px;margin:0;">${content.artist || ''}</h2>`;
    html += `<h3 style="font-family:var(--font-body);font-size:16pt;color:#9ca3af;font-style:italic;font-weight:300;text-transform:lowercase;line-height:1.1;margin:0;">${content.name || ''}</h3>`;
    if (PARAMS.showPrice && val(content.price)) {
        html += `<div style="font-family:var(--font-body);font-size:10pt;letter-spacing:-0.5px;color:#000;font-style:italic;">${content.price}</div>`;
    }
    html += '</div>';

    // Meta grid
    let metaItems = '';
    if (val(content.category)) metaItems += `<div><span style="display:block;font-family:\'Capriola\',sans-serif;font-size:8px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">Category</span><span style="font-family:var(--font-body);font-size:8.5pt;font-style:italic;">${content.category}</span></div>`;
    if (val(content.type)) metaItems += `<div><span style="display:block;font-family:\'Capriola\',sans-serif;font-size:8px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">Product Type</span><span style="font-family:var(--font-body);font-size:8.5pt;font-style:italic;">${content.type}</span></div>`;
    if (val(content.color)) metaItems += `<div><span style="display:block;font-family:\'Capriola\',sans-serif;font-size:8px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">Color</span><span style="font-family:var(--font-body);font-size:8.5pt;font-style:italic;">${content.color}</span></div>`;
    if (val(content.edition)) metaItems += `<div><span style="display:block;font-family:\'Capriola\',sans-serif;font-size:8px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">Edition</span><span style="font-family:var(--font-body);font-size:8.5pt;font-style:italic;">${content.edition}</span></div>`;
    if (val(content.dimensions)) metaItems += `<div><span style="display:block;font-family:\'Capriola\',sans-serif;font-size:8px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">Dimensions</span><span style="font-family:var(--font-body);font-size:8.5pt;font-style:italic;">${content.dimensions}</span></div>`;
    if (metaItems) {
        html += `<div style="display:grid;grid-template-columns:1fr 1fr;gap:1.5rem 2rem;border-top:1px solid rgba(0,0,0,0.08);border-bottom:1px solid rgba(0,0,0,0.08);padding:2rem 0;">${metaItems}</div>`;
    }

    // Description
    if (val(content.description)) {
        html += '<div style="margin-top:2.5rem;">';
        html += `<h4 style="font-family:var(--font-headers);font-size:9px;text-transform:uppercase;font-weight:bold;letter-spacing:0.2em;color:#000;margin-bottom:0.5rem;">Product Description</h4>`;
        html += `<p style="font-family:var(--font-body);font-size:8.5pt;color:#666;font-style:italic;font-weight:300;text-align:justify;text-align-last:left;line-height:1.7;white-space:pre-line;">${content.description}</p>`;
        html += '</div>';
    }

    // Artist Bio
    if (val(content.bio)) {
        html += '<div style="margin-top:2.5rem;padding-top:2.5rem;border-top:1px solid rgba(0,0,0,0.08);">';
        html += `<h4 style="font-family:\'Capriola\',sans-serif;font-size:9px;text-transform:uppercase;letter-spacing:0.2em;color:#9ca3af;margin-bottom:0.5rem;">About ${content.artist || ''}</h4>`;
        html += `<p style="font-family:var(--font-body);font-size:8pt;color:#9ca3af;font-style:italic;letter-spacing:-0.3px;font-weight:300;text-align:justify;text-align-last:left;line-height:1.6;white-space:pre-line;">${content.bio}</p>`;
        html += '</div>';
    }

    document.getElementById('gm-details').innerHTML = html;
    document.getElementById('grid-modal-overlay').classList.add('open');

    // Reset scroll
    const infoPanel = document.querySelector('.gm-info-panel');
    if (infoPanel) infoPanel.scrollTop = 0;
    const modal = document.getElementById('grid-modal');
    if (modal) modal.scrollTop = 0;

    // Swipe support for mobile carousel
    const imgWrap = document.getElementById('gm-img-wrap');
    if (imgWrap && !imgWrap._gmSwipeInit) {
        imgWrap._gmSwipeInit = true;
        let startX = 0, endX = 0;
        imgWrap.addEventListener('touchstart', (e) => { startX = e.changedTouches[0].screenX; }, { passive: true });
        imgWrap.addEventListener('touchend', (e) => {
            endX = e.changedTouches[0].screenX;
            const diff = startX - endX;
            if (Math.abs(diff) > 50) {
                if (diff > 0) window.gmCarouselNext();
                else window.gmCarouselPrev();
            }
        }, { passive: true });
    }
};

window.closeGridModal = function() {
    document.getElementById('grid-modal-overlay').classList.remove('open');
};

// Escape key closes grid modal
window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && document.getElementById('grid-modal-overlay').classList.contains('open')) {
        window.closeGridModal();
    }
});

window.toggleInfoPanel = function(panelType) {
    const infoPanel = document.getElementById('info-panel');
    const btnStatus = document.getElementById('btn-status');
    const btnAbout = document.getElementById('btn-about');
    const panelTitle = document.getElementById('info-panel-title');
    const isMobileNow = isMobileView();
    if (currentInfoPanel === panelType && infoPanel.classList.contains('open')) {
        closeInfoPanel();
        return;
    }
    if (isMobileNow) {
        if (document.getElementById('side-panel').classList.contains('open')) {
            closePanel();
        }
    }
    currentInfoPanel = panelType;
    if (btnStatus) btnStatus.classList.toggle('active', panelType === 'status');
    if (btnAbout) btnAbout.classList.toggle('active', panelType === 'about');
    const content = document.getElementById('info-panel-content');
    if (panelType === 'status') {
        panelTitle.textContent = 'STATUS CO';
        content.innerHTML = `
            <img class="info-panel-logo" src="https://raw.githubusercontent.com/decentralize-dfw/co-contemporaryartifacts-glb/main/logo_text.png" alt="Status Co">
            <div class="info-panel-text">${infoTextStatus}</div>
        `;
    } else {
        panelTitle.textContent = 'CONTEMPORARY ARTIFACTS';
        content.innerHTML = `
            <div class="info-panel-text">${infoTextAbout}</div>
        `;
    }
    infoPanel.classList.add('open');
};
window.closeInfoPanel = function() {
    const infoPanel = document.getElementById('info-panel');
    const btnStatus = document.getElementById('btn-status');
    const btnAbout = document.getElementById('btn-about');
    infoPanel.classList.remove('open');
    if (btnStatus) btnStatus.classList.remove('active');
    if (btnAbout) btnAbout.classList.remove('active');
    currentInfoPanel = null;
};
window.toggleNavMenu = function() {
    const dropdown = document.getElementById('nav-dropdown');
    if (!dropdown) return;
    dropdown.classList.toggle('open');
};
function updateOrthoFrustum() {
    if (!cameraOrtho) return;
    const aspect = window.innerWidth / window.innerHeight;
    const box = new THREE.Box3();
    clickableObjects.forEach(o => box.expandByObject(o));
    const size = box.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z) * 1.5;
    cameraOrtho.left = -maxDim * aspect; 
    cameraOrtho.right = maxDim * aspect;
    cameraOrtho.top = maxDim; 
    cameraOrtho.bottom = -maxDim;
    cameraOrtho.updateProjectionMatrix();
}
let marqueeAnimId = null;
function initMarquee(text) {
    const track = document.getElementById('marquee-track');
    if (!track || !text) return;
    track.innerHTML = '';
    const measure = document.createElement('span');
    measure.className = 'marquee-content';
    measure.textContent = text;
    measure.style.visibility = 'hidden';
    measure.style.position = 'absolute';
    document.body.appendChild(measure);
    const singleW = measure.offsetWidth;
    document.body.removeChild(measure);
    const copies = Math.ceil((window.innerWidth * 3) / Math.max(singleW, 1)) + 2;
    for (let i = 0; i < copies; i++) {
        const span = document.createElement('span');
        span.className = 'marquee-content';
        span.textContent = text;
        track.appendChild(span);
    }
    positionMarquee();
    startMarqueeAnimation();
}
function startMarqueeAnimation() {
    if (marqueeAnimId) cancelAnimationFrame(marqueeAnimId);
    const track = document.getElementById('marquee-track');
    if (!track || !track.children.length) return;
    const firstSpan = track.children[0];
    let offset = 0;
    const getSpeed = () => {
        const isMobileNow = isMobileView();
        const prop = isMobileNow ? '--marquee-mobile-speed' : '--marquee-speed';
        const s = parseFloat(getComputedStyle(document.documentElement).getPropertyValue(prop)) || 20;
        return 80 / s;
    };
    function tick() {
        offset -= getSpeed();
        const singleW = firstSpan.offsetWidth;
        if (singleW > 0 && Math.abs(offset) >= singleW) {
            offset += singleW;
        }
        track.style.transform = `translateX(${offset}px)`;
        marqueeAnimId = requestAnimationFrame(tick);
    }
    marqueeAnimId = requestAnimationFrame(tick);
}
window.mobileMarqueeEnabled = getParam('--marquee-mobile-enabled') === 'yes';
window.positionMarquee = positionMarquee;
function positionMarquee() {
    positionRightToggles();
    const bar = document.getElementById('marquee-bar');
    if (!bar) return;
    const isMobileNow = isMobileView();
    const infoButtons = document.getElementById('info-buttons');
    if (!infoButtons) return;
    if (isMobileNow && !mobileMarqueeEnabled) {
        bar.style.display = 'none';
        return;
    }
    bar.style.display = 'flex';
    const leftEnd = infoButtons.getBoundingClientRect().right + 15;
    let rightEnd;
    if (isMobileNow) {
        const mobileTopOffset = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-top-offset')) || 12;
        bar.style.top = mobileTopOffset + 'px';
        rightEnd = window.innerWidth - 12;
    } else {
        const row0 = document.getElementById('right-toggle-row0');
        const row1 = document.getElementById('right-toggle-row1');
        const row2 = document.getElementById('right-toggle-row2');
        let minLeft = window.innerWidth - 80;
        [row0, row1, row2].forEach(r => {
            if (r && r.style.display !== 'none') {
                const rect = r.getBoundingClientRect();
                if (rect.width > 0) minLeft = Math.min(minLeft, rect.left);
            }
        });
        rightEnd = minLeft - 15;
    }
    const width = rightEnd - leftEnd;
    if (width < 50) { bar.style.display = 'none'; return; }
    bar.style.left = leftEnd + 'px';
    bar.style.width = width + 'px';
    bar.style.transform = 'none';
}
function onWindowResize() {
    markDirty();
    const aspect = window.innerWidth / window.innerHeight;
    camera.aspect = aspect;
    camera.updateProjectionMatrix();
    updateOrthoFrustum();
    renderer.setSize(window.innerWidth, window.innerHeight);
    composer.setSize(window.innerWidth, window.innerHeight);
    positionMarquee();

    const isMobileNow = isMobileView();
    if (isMobileNow !== wasMobileView) {
        wasMobileView = isMobileNow;
        // Switch camera: mobile camera for mobile, cam 1 for desktop
        if (isMobileNow) {
            if (currentCameraIndex !== 'mobile') {
                switchCamera('mobile');
            }
        } else {
            if (currentCameraIndex === 'mobile') {
                switchCamera(1);
            }
        }
        // Switch HDRI preset for the new mode
        const targetHdri = isMobileNow ? (runtimeParams.hdriPresetMobile || 1) : (runtimeParams.hdriPreset || 1);
        const currentActiveHdri = isMobileNow ? currentHdriPresetMobile : currentHdriPreset;
        if (targetHdri !== currentActiveHdri) {
            loadHDRI(targetHdri, null, isMobileNow);
        }
        // Re-apply all scene/env settings for the new mobile/desktop state
        applyRuntimeParameters();
    }
}
function getLocalLockMode() {
    if (currentCameraIndex === 1) return runtimeParams.cam1LockMode || 'none';
    if (currentCameraIndex === 2) return runtimeParams.cam2LockMode || 'none';
    if (currentCameraIndex === 3) return runtimeParams.cam3LockMode || 'none';
    if (currentCameraIndex === 'mobile') return runtimeParams.camMobileLockMode || 'none';
    return 'none';
}
function getEffectiveLockMode() {
    if (isCameraLocked) return 'full';
    return getLocalLockMode();
}
function isEffectivelyLocked() {
    return getEffectiveLockMode() !== 'none';
}
window.applyLockState = applyLockState;
function applyLockState() {
    markDirty();
    const mode = getEffectiveLockMode();
    const activeControls = currentCameraMode === 'persp' ? controls : controlsOrtho;
    const inactiveControls = currentCameraMode === 'persp' ? controlsOrtho : controls;
    const isMobileNow = isMobileView();

    // Always disable the inactive controls
    inactiveControls.enabled = false;

    // Always enable the active controls container - individual features are controlled below
    activeControls.enabled = true;

    // Polar angle conversion: user params are elevation (0=horizon, 90=top)
    // Three.js polar is from zenith (0=top, PI=bottom): (90 - elevation) * PI/180
    const polarMin = (90 - runtimeParams.cameraMaxPolarAngle) * Math.PI / 180;
    const polarMax = (90 - runtimeParams.cameraMinPolarAngle) * Math.PI / 180;

    // ZOOM-ROTATE: When zoomed into an object with zoom-rotate enabled
    // If camera is locked → only horizontal rotation (freeze polar angle)
    // If camera is unlocked → full rotation within configured limits
    if (currentSelectedObject && runtimeParams.zoomRotateEnabled) {
        activeControls.enableRotate = true;
        activeControls.enableZoom = true;
        activeControls.enablePan = false;
        if (mode !== 'none') {
            const currentPolar = activeControls.getPolarAngle();
            activeControls.minPolarAngle = currentPolar;
            activeControls.maxPolarAngle = currentPolar;
        } else {
            activeControls.minPolarAngle = polarMin;
            activeControls.maxPolarAngle = polarMax;
        }
        return;
    }

    // FULL LOCK: No rotation, no pan, no zoom
    if (mode === 'full') {
        activeControls.enableRotate = false;
        activeControls.enablePan = false;
        activeControls.enableZoom = false;
        if (camDotFadeTimer) {
            clearTimeout(camDotFadeTimer);
            camDotFadeTimer = null;
        }
        if (activeCamDot) {
            const dot = document.getElementById('cam-dot-' + activeCamDot);
            if (dot) {
                dot.classList.remove('fading');
                dot.classList.add('active');
            }
        }
        return;
    }

    // HORIZONTAL LOCK: Only azimuthal rotation (around Y), zoom allowed, no pan
    if (mode === 'horizontal') {
        const currentPolar = activeControls.getPolarAngle();
        activeControls.minPolarAngle = currentPolar;
        activeControls.maxPolarAngle = currentPolar;
        activeControls.enableRotate = true;
        activeControls.enablePan = false;
        activeControls.enableZoom = true;
        return;
    }

    // NONE (unlocked): Full freedom with configured limits
    activeControls.enableRotate = true;
    activeControls.enablePan = isMobileNow && PARAMS.mobilePanEnabled;
    activeControls.enableZoom = true;
    activeControls.minPolarAngle = polarMin;
    activeControls.maxPolarAngle = polarMax;
}
window.toggleCameraLock = function() {
    isCameraLocked = !isCameraLocked;
    const lockedPath = 'M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z';
    const unlockedPath = 'M12 17c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm6-9h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6h1.9c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm0 12H6V10h12v10z';
    const lockBtns = [document.getElementById('btn-lock-icon'), document.getElementById('btn-lock-icon-public')];
    lockBtns.forEach(btn => {
        if (!btn) return;
        const path = btn.querySelector('path');
        if (path) path.setAttribute('d', isCameraLocked ? lockedPath : unlockedPath);
        btn.style.opacity = isCameraLocked ? '1' : '0.5';
    });
    applyLockState();
};
function applyInitialLockState() {
    const lockBtn = document.getElementById('btn-lock-icon-public');
    if (lockBtn) {
        lockBtn.style.display = runtimeParams.cameraLockShowPublic ? '' : 'none';
    }
    if (isCameraLocked) {
        const lockedPath = 'M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z';
        const lockBtns = [document.getElementById('btn-lock-icon'), document.getElementById('btn-lock-icon-public')];
        lockBtns.forEach(btn => {
            if (!btn) return;
            const path = btn.querySelector('path');
            if (path) path.setAttribute('d', lockedPath);
            btn.style.opacity = '1';
        });
        applyLockState();
    } else if (getLocalLockMode() !== 'none') {
        applyLockState();
    }
}
function animate() {
    requestAnimationFrame(animate);
    // Controls must update every frame for damping
    if (currentCameraMode === 'persp') {
        controls.update();
    } else {
        controlsOrtho.update();
    }
    // Continuous animations keep scene dirty
    const panelOpen = document.getElementById('side-panel').classList.contains('open');
    const hasSelectedRotation = currentSelectedObject && !editorModeActive && !isMobileView() &&
        (PARAMS.selectedRotateX !== 0 || PARAMS.selectedRotateY !== 0 || PARAMS.selectedRotateZ !== 0);
    if (hasSelectedRotation || currentHoveredObject || activeTweenCount > 0) {
        sceneDirty = true;
    }
    if (!sceneDirty) return;
    sceneDirty = false;
    if (currentCameraMode === 'persp') {
        composer.passes[0].camera = camera;
        if (ssaoPass) ssaoPass.camera = camera;
    } else {
        composer.passes[0].camera = cameraOrtho;
        if (ssaoPass) ssaoPass.camera = cameraOrtho;
    }
    if(currentHoveredObject) {
        const bboxTop = currentHoveredObject.userData._bboxTopLocal || 0;
        const f = currentHoveredObject.userData._hoverLiftMult || 1.0;
        const liftDist = PARAMS.hoverLiftDistance * f;
        const baseY = currentHoveredObject.position.y + bboxTop + liftDist;
        if(hoverLabelArtist) {
            const artistY = baseY + PARAMS.hoverLabelOffsetY * f;
            hoverLabelArtist.position.set(currentHoveredObject.position.x, artistY, currentHoveredObject.position.z);
        }
        if(hoverLabel) {
            const nameY = baseY + (PARAMS.hoverLabelOffsetY + PARAMS.hoverLabelNameOffsetY) * f;
            hoverLabel.position.set(currentHoveredObject.position.x, nameY, currentHoveredObject.position.z);
        }
    }
    if(hasSelectedRotation) {
        const dt = 0.016;
        currentSelectedObject.rotation.x += PARAMS.selectedRotateX * dt;
        currentSelectedObject.rotation.y += PARAMS.selectedRotateY * dt;
        currentSelectedObject.rotation.z += PARAMS.selectedRotateZ * dt;
    }
    composer.render();
}
function TWEEN(obj) {
    this.target = obj;
    this.to = (dest, dur) => { this.dest = dest; this.dur = dur; return this; };
    this.easing = (fn) => { this.ease = fn; return this; };
    this.onUpdate = (fn) => { this.updateFn = fn; return this; };
    this.onComplete = (fn) => { this.completeFn = fn; return this; };
    this.start = () => {
        const start = {}; for(let k in this.dest) start[k] = this.target[k];
        // Handle zero/tiny duration: set values immediately
        if (!this.dur || this.dur <= 0) {
            for(let k in this.dest) this.target[k] = this.dest[k];
            if(this.updateFn) this.updateFn();
            if(this.completeFn) this.completeFn();
            markDirty();
            return;
        }
        const startTime = Date.now();
        const ease = this.ease || (t => t*(2-t));
        activeTweenCount++;
        markDirty();
        const tick = () => {
            const now = Date.now();
            let p = (now - startTime) / this.dur;
            if(p > 1) p = 1;
            const e = ease(p);
            for(let k in this.dest) this.target[k] = start[k]+(this.dest[k]-start[k])*e;
            if(this.updateFn) this.updateFn();
            markDirty();
            if(p < 1) {
                requestAnimationFrame(tick);
            } else {
                activeTweenCount--;
                if(this.completeFn) this.completeFn();
            }
        }; tick();
    };
}
function buildBoundingBoxes() {
    boundingBoxHelpers.forEach(h => scene.remove(h));
    boundingBoxHelpers = [];
    clickableObjects.forEach(obj => {
        if (!obj.userData.clickable) return;
        const box = new THREE.Box3().setFromObject(obj);
        const helper = new THREE.Box3Helper(box, 0x000000);
        helper.material.transparent = true;
        helper.material.opacity = 0.3;
        helper.material.linewidth = 1;
        helper.visible = boundingBoxesVisible;
        helper.userData.isBoundingBoxHelper = true;
        scene.add(helper);
        boundingBoxHelpers.push(helper);
    });
}
function buildClickableZones() {
    clickableZones.forEach(z => scene.remove(z));
    clickableZones.length = 0;
    const isMobileNow = isMobileView();
    clickableObjects.forEach(obj => {
        if (!obj.userData.clickable) return;
        const box = new THREE.Box3().setFromObject(obj);
        const size = box.getSize(new THREE.Vector3());
        const center = box.getCenter(new THREE.Vector3());
        const geometry = new THREE.BoxGeometry(
            Math.max(size.x, 50),
            size.y,
            Math.max(size.z, 50)
        );
        const material = new THREE.MeshBasicMaterial({
            visible: false,
            side: THREE.DoubleSide
        });
        const zoneMesh = new THREE.Mesh(geometry, material);
        zoneMesh.position.set(center.x, center.y, center.z);
        zoneMesh.userData.parentObject = obj;
        zoneMesh.userData.isClickableZone = true;
        zoneMesh.raycast = THREE.Mesh.prototype.raycast;
        scene.add(zoneMesh);
        clickableZones.push(zoneMesh);
    });
}
function toggleBoundingBoxes() {
    boundingBoxesVisible = !boundingBoxesVisible;
    boundingBoxHelpers.forEach(h => {
        h.visible = boundingBoxesVisible;
    });
}
function updateClickableZonePosition(obj) {
    const zone = clickableZones.find(z => z.userData.parentObject === obj);
    if (!zone) return;
    const origPos = obj.userData.originalPosition;
    const currentPos = obj.position.clone();
    obj.position.copy(origPos);
    const box = new THREE.Box3().setFromObject(obj);
    obj.position.copy(currentPos);
    const size = box.getSize(new THREE.Vector3());
    const center = box.getCenter(new THREE.Vector3());
    zone.geometry.dispose();
    zone.geometry = new THREE.BoxGeometry(
        Math.max(size.x, 50),
        size.y,
        Math.max(size.z, 50)
    );
    zone.position.set(center.x, center.y, center.z);
}
function showMobileLabels() {
    hideMobileLabels();
    mobileLabelsVisible = true;
    const maxW = PARAMS.hoverLabelFontBox;
    const rotX = PARAMS.hoverLabelRotationX * Math.PI / 180;
    const rotY = PARAMS.hoverLabelRotationY * Math.PI / 180;
    const rotZ = PARAMS.hoverLabelRotationZ * Math.PI / 180;
    clickableObjects.forEach(obj => {
        if (!obj.userData.clickable) return;
        const content = obj.userData.modalContent;
        if (!content) return;
        const box = new THREE.Box3().setFromObject(obj);
        const bboxTopLocal = box.max.y - obj.position.y;
        const artistText = content.artist || '';
        if (artistText) {
            const label = createHoverTextMesh(artistText, PARAMS.hoverArtistFontSize, maxW, {
                fontStyle: PARAMS.hoverArtistFontStyle,
                letterSpacing: PARAMS.hoverArtistLetterSpacing,
                color: PARAMS.hoverArtistColor,
                opacity: PARAMS.hoverArtistOpacity
            });
            const y = obj.position.y + bboxTopLocal + PARAMS.hoverLabelOffsetY;
            label.position.set(obj.position.x, y, obj.position.z);
            label.rotation.set(rotX, rotY, rotZ);
            scene.add(label);
            mobileStaticLabels.push(label);
        }
        const nameText = content.hoverDisplayName || content.name || '';
        if (nameText) {
            const label = createHoverTextMesh(nameText, PARAMS.hoverLabelFontSize, maxW, {
                fontStyle: PARAMS.hoverLabelFontStyle,
                letterSpacing: PARAMS.hoverLabelLetterSpacing,
                color: PARAMS.hoverLabelColor,
                opacity: PARAMS.hoverLabelOpacity
            });
            const y = obj.position.y + bboxTopLocal + PARAMS.hoverLabelOffsetY + PARAMS.hoverLabelNameOffsetY;
            label.position.set(obj.position.x, y, obj.position.z);
            label.rotation.set(rotX, rotY, rotZ);
            scene.add(label);
            mobileStaticLabels.push(label);
        }
    });
}
function hideMobileLabels() {
    mobileLabelsVisible = false;
    mobileStaticLabels.forEach(l => scene.remove(l));
    mobileStaticLabels = [];
}
window.toggleEditorMode = function() {
    editorModeActive = !editorModeActive;
    const btnEditor = document.getElementById('btn-editor');
    const editorPanel = document.getElementById('editor-panel');
    const transformGroup = document.getElementById('transform-group');
    const btnCam3 = document.getElementById('btn-cam3');
    const btnCamMobile = document.getElementById('btn-cam-mobile');
    const uiContainer = document.getElementById('ui-container');
    if (editorModeActive) {
        if(btnEditor) btnEditor.classList.add('active');
        editorPanel.classList.add('open');
        if(uiContainer) uiContainer.classList.add('editor-visible');
        transformGroup.style.display = 'flex';
        transformControls.visible = true;
        transformControls.enabled = true;
        if(btnCam3) btnCam3.style.display = 'inline-block';
        if(btnCamMobile) btnCamMobile.style.display = 'inline-block';
        if (isMobileView() && currentCameraIndex !== 'mobile') {
            switchCamera('mobile');
        }
        if (!selectedEditorObject && clickableObjects.length > 0) {
            for (let i = 0; i < clickableObjects.length; i++) {
                const idx = clickableObjects[i].userData.sceneObjectIndex;
                if (!lockedObjects.has(idx)) {
                    selectEditorObject(clickableObjects[i]);
                    break;
                }
            }
        }
    } else {
        if(btnEditor) btnEditor.classList.remove('active');
        editorPanel.classList.remove('open');
        if(uiContainer) uiContainer.classList.remove('editor-visible');
        transformGroup.style.display = 'none';
        transformControls.visible = false;
        transformControls.enabled = false;
        transformControls.detach();
        selectedEditorObject = null;
        updateCam3Visibility();
        if(btnCamMobile) btnCamMobile.style.display = 'none';
        if(currentCameraMode === 'persp' || currentCameraIndex === 'mobile') {
            switchCamera(isMobileView() ? 'mobile' : 1);
        }
    }
};
window.setTransformMode = function(mode) {
    currentTransformMode = mode;
    transformControls.setMode(mode);
    document.getElementById('btn-translate').classList.toggle('active', mode === 'translate');
    document.getElementById('btn-rotate').classList.toggle('active', mode === 'rotate');
    document.getElementById('btn-scale').classList.toggle('active', mode === 'scale');
};
(function initEditorDrag() {
    const panel = document.getElementById('editor-panel');
    const handle = document.getElementById('editor-drag-handle');
    if (!handle || !panel) return;
    let isDragging = false, startX, startY, startLeft, startTop;
    handle.addEventListener('mousedown', (e) => {
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        startLeft = panel.offsetLeft;
        startTop = panel.offsetTop;
        e.preventDefault();
    });
    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        panel.style.left = (startLeft + e.clientX - startX) + 'px';
        panel.style.top = (startTop + e.clientY - startY) + 'px';
    });
    document.addEventListener('mouseup', () => { isDragging = false; });
})();
function updateCameraParamsDisplay() {
    const activeCamera = getActiveCamera();
    document.getElementById('cam-pos-x').value = activeCamera.position.x.toFixed(0);
    document.getElementById('cam-pos-y').value = activeCamera.position.y.toFixed(0);
    document.getElementById('cam-pos-z').value = activeCamera.position.z.toFixed(0);
    if (activeCamera.isPerspectiveCamera) {
        document.getElementById('cam-fov').value = activeCamera.fov.toFixed(0);
    }
}
window.updateCameraFOV = function(value) {
    const fov = parseFloat(value) || 60;
    const activeCamera = getActiveCamera();
    if (activeCamera.isPerspectiveCamera) {
        activeCamera.fov = fov;
        activeCamera.updateProjectionMatrix();
    }
};
window.updateCameraPosition = function() {
    const x = parseFloat(document.getElementById('cam-pos-x').value) || 0;
    const y = parseFloat(document.getElementById('cam-pos-y').value) || 0;
    const z = parseFloat(document.getElementById('cam-pos-z').value) || 0;
    const activeCamera = getActiveCamera();
    activeCamera.position.set(x, y, z);
    activeCamera.updateProjectionMatrix();
};
function buildEditorPanel() {
    const tbody = document.getElementById('editor-table-body');
    tbody.innerHTML = '';
    sceneObjects.forEach((item, index) => {
        const obj = loadedObjectsMap.get(index);
        if (!obj) return;
        const name = item.modalContent?.name || `Object ${index + 1}`;
        const pos = obj.position;
        const rot = obj.rotation;
        const scl = obj.scale.x;
        const isLocked = lockedObjects.has(index);
        const tr = document.createElement('tr');
        tr.dataset.index = index;
        if (isLocked) tr.classList.add('locked');
        tr.onclick = (e) => {
            if (e.target.tagName !== 'INPUT') selectEditorObjectByIndex(index);
        };
        tr.innerHTML = `
            <td class="lock-cell"><span class="lock-icon" onclick="toggleLock(${index}, event)">${isLocked ? '🔒' : '🔓'}</span></td>
            <td class="name-cell" title="${name}">${name}</td>
            <td class="col-x"><input type="number" data-prop="posX" data-index="${index}" value="${pos.x.toFixed(0)}" step="10" ${isLocked ? 'disabled' : ''}></td>
            <td class="col-y"><input type="number" data-prop="posY" data-index="${index}" value="${pos.y.toFixed(0)}" step="10" ${isLocked ? 'disabled' : ''}></td>
            <td class="col-z"><input type="number" data-prop="posZ" data-index="${index}" value="${pos.z.toFixed(0)}" step="10" ${isLocked ? 'disabled' : ''}></td>
            <td class="col-x"><input type="number" data-prop="rotX" data-index="${index}" value="${(rot.x * 180 / Math.PI).toFixed(0)}" step="5" ${isLocked ? 'disabled' : ''}></td>
            <td class="col-y"><input type="number" data-prop="rotY" data-index="${index}" value="${(rot.y * 180 / Math.PI).toFixed(0)}" step="5" ${isLocked ? 'disabled' : ''}></td>
            <td class="col-z"><input type="number" data-prop="rotZ" data-index="${index}" value="${(rot.z * 180 / Math.PI).toFixed(0)}" step="5" ${isLocked ? 'disabled' : ''}></td>
            <td class="col-s"><input type="number" data-prop="scale" data-index="${index}" value="${scl.toFixed(0)}" step="10" ${isLocked ? 'disabled' : ''}></td>
        `;
        tbody.appendChild(tr);
    });
    tbody.querySelectorAll('input').forEach(input => {
        input.addEventListener('focus', (e) => {
            saveUndoState();
            e.target.select();
        });
        input.addEventListener('change', handleEditorInputChange);
        input.addEventListener('input', handleEditorInputChange);
        input.addEventListener('click', (e) => e.stopPropagation());
        input.addEventListener('mousedown', (e) => e.stopPropagation());
        input.addEventListener('keydown', (e) => e.stopPropagation());
    });
}
window.toggleLock = function(index, event) {
    event.stopPropagation();
    if (lockedObjects.has(index)) lockedObjects.delete(index);
    else {
        lockedObjects.add(index);
        if (selectedEditorObject && selectedEditorObject.userData.sceneObjectIndex === index) {
            transformControls.detach();
            selectedEditorObject = null;
        }
    }
    buildEditorPanel();
    if (selectedEditorObject) selectEditorObject(selectedEditorObject);
};
function handleEditorInputChange(e) {
    const input = e.target;
    const index = parseInt(input.dataset.index);
    const prop = input.dataset.prop;
    const value = parseFloat(input.value) || 0;
    const obj = loadedObjectsMap.get(index);
    if (!obj) return;
    switch(prop) {
        case 'posX': obj.position.x = value; break;
        case 'posY': obj.position.y = value; break;
        case 'posZ': obj.position.z = value; break;
        case 'rotX': obj.rotation.x = value * Math.PI / 180; break;
        case 'rotY': obj.rotation.y = value * Math.PI / 180; break;
        case 'rotZ': obj.rotation.z = value * Math.PI / 180; break;
        case 'scale': obj.scale.setScalar(value); break;
    }
    obj.userData.originalPosition = obj.position.clone();
}
function selectEditorObjectByIndex(index) {
    if (lockedObjects.has(index)) return;
    const obj = loadedObjectsMap.get(index);
    if (obj) selectEditorObject(obj);
}
function selectEditorObject(obj) {
    if (lockedObjects.has(obj.userData.sceneObjectIndex)) return;
    if (!obj.parent) return;
    selectedEditorObject = obj;
    document.querySelectorAll('#editor-table-body tr').forEach(row => {
        row.classList.remove('selected');
        if (parseInt(row.dataset.index) === obj.userData.sceneObjectIndex) {
            row.classList.add('selected');
        }
    });
    if (transformControls && editorModeActive) transformControls.attach(obj);
}
function updateEditorPanelValues() {
    if (!selectedEditorObject) return;
    const index = selectedEditorObject.userData.sceneObjectIndex;
    const pos = selectedEditorObject.position;
    const rot = selectedEditorObject.rotation;
    const scl = selectedEditorObject.scale.x;
    const inputs = document.querySelectorAll(`#editor-table-body input[data-index="${index}"]`);
    inputs.forEach(input => {
        const prop = input.dataset.prop;
        let value;
        switch(prop) {
            case 'posX': value = pos.x.toFixed(0); break;
            case 'posY': value = pos.y.toFixed(0); break;
            case 'posZ': value = pos.z.toFixed(0); break;
            case 'rotX': value = (rot.x * 180 / Math.PI).toFixed(0); break;
            case 'rotY': value = (rot.y * 180 / Math.PI).toFixed(0); break;
            case 'rotZ': value = (rot.z * 180 / Math.PI).toFixed(0); break;
            case 'scale': value = scl.toFixed(0); break;
        }
        if (document.activeElement !== input) input.value = value;
    });
    selectedEditorObject.userData.originalPosition = selectedEditorObject.position.clone();
}
window.switchCamera = function(mode) {
    const btnCam3p = document.getElementById('btn-cam3');
    const btnCam1 = document.getElementById('btn-cam1');
    const btnCam2 = document.getElementById('btn-cam2');
    const btnCamMobile = document.getElementById('btn-cam-mobile');
    const btnLockIcon = document.getElementById('btn-lock-icon');
    [btnCam3p, btnCam1, btnCam2, btnCamMobile].forEach(b => { if(b) b.classList.remove('active') });

    // If mobile was in persp mode, restore Camera 3's perspective camera state
    if (savedCam3State) {
        camera.position.copy(savedCam3State.position);
        camera.fov = savedCam3State.fov;
        camera.updateProjectionMatrix();
        controls.target.copy(savedCam3State.target);
        controls.update();
        savedCam3State = null;
    }

    camSwitching = true;
    if (camDotFadeTimer) { clearTimeout(camDotFadeTimer); camDotFadeTimer = null; }
    const dot1 = document.getElementById('cam-dot-1');
    const dot2 = document.getElementById('cam-dot-2');
    const dot3 = document.getElementById('cam-dot-3');
    if (dot1) dot1.classList.remove('active', 'fading');
    if (dot2) dot2.classList.remove('active', 'fading');
    if (dot3) dot3.classList.remove('active', 'fading');
    activeCamDot = null;
    if (mode === 1 || mode === 2) {
        const dot = document.getElementById('cam-dot-' + mode);
        if (dot) dot.classList.add('active');
        activeCamDot = mode;
    } else if (mode === 'persp' && runtimeParams.cam3ShowPublic) {
        if (dot3) dot3.classList.add('active');
        activeCamDot = 3;
    }
    if (mode === 'mobile') {
        currentCameraMode = 'ortho';
        currentCameraIndex = 'mobile';
        if(btnCamMobile) btnCamMobile.classList.add('active');
        controls.enabled = false;
        controlsOrtho.enabled = false;
        if (transformControls) transformControls.camera = cameraOrtho;
        // Pre-set perspective camera FOV from parameter so it's ready for persp toggle
        camera.fov = runtimeParams.camMobileFov || 45;
        camera.updateProjectionMatrix();
        updateOrthoFrustum();
        const tx = runtimeParams.camMobilePosX;
        const ty = runtimeParams.camMobilePosY;
        const tz = runtimeParams.camMobilePosZ;
        const targetUp = new THREE.Vector3(0, 1, 0);
        const startPos = cameraOrtho.position.clone();
        const startRot = cameraOrtho.quaternion.clone();
        const startUp = cameraOrtho.up.clone();
        const endPos = new THREE.Vector3(tx, ty, tz);
        const dummyCam = cameraOrtho.clone();
        dummyCam.position.copy(endPos);
        dummyCam.up.copy(targetUp);
        dummyCam.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        const endRot = dummyCam.quaternion.clone();
        const mobileFovForZoom = runtimeParams.camMobileFov || 45;
        cameraOrtho.zoom = runtimeParams.initialCameraZoomMobile * (45 / mobileFovForZoom);
        cameraOrtho.updateProjectionMatrix();
        const animData = { t: 0 };
        new TWEEN(animData)
            .to({ t: 1 }, 2000)
            .easing(t => t)
            .onUpdate(() => {
                const k = animData.t;
                cameraOrtho.position.lerpVectors(startPos, endPos, k);
                cameraOrtho.up.lerpVectors(startUp, targetUp, k).normalize();
                cameraOrtho.quaternion.slerpQuaternions(startRot, endRot, k);
            })
            .onComplete(() => {
                camSwitching = false;
                cameraOrtho.position.copy(endPos);
                cameraOrtho.up.copy(targetUp);
                cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.update();
                applyLockState();
            })
            .start();
    } else if (mode === 'persp') {
        currentCameraMode = 'persp';
        currentCameraIndex = 3;
        if(btnCam3p) btnCam3p.classList.add('active');
        controlsOrtho.enabled = false;
        if (transformControls) transformControls.camera = camera;
        camera.position.set(runtimeParams.cam3PosX, runtimeParams.cam3PosY, runtimeParams.cam3PosZ);
        camera.fov = PARAMS.cameraPerspectiveFov;
        camera.updateProjectionMatrix();
        camera.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controls.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controls.update();
        camSwitching = false;
        applyLockState();
    } else {
        currentCameraMode = 'ortho';
        currentCameraIndex = mode;
        if(mode === 1) btnCam1.classList.add('active');
        if(mode === 2) btnCam2.classList.add('active');
        controls.enabled = false;
        controlsOrtho.enabled = false;
        if (transformControls) transformControls.camera = cameraOrtho;
        updateOrthoFrustum();
        let tx=0, ty=0, tz=0;
        let targetUp = new THREE.Vector3(0, 1, 0);
        if(mode === 1) {
            tx = runtimeParams.cam1PosX; ty = runtimeParams.cam1PosY; tz = runtimeParams.cam1PosZ;
        }
        else if(mode === 2) {
            tx = runtimeParams.cam2PosX; ty = runtimeParams.cam2PosY; tz = runtimeParams.cam2PosZ;
        }
        const orthoFovCam12 = runtimeParams.cameraOrthoFov || 45;
        cameraOrtho.zoom = runtimeParams.initialCameraZoom * (45 / orthoFovCam12);
        cameraOrtho.updateProjectionMatrix();
        const startPos = cameraOrtho.position.clone();
        const startRot = cameraOrtho.quaternion.clone();
        const startUp = cameraOrtho.up.clone();
        const endPos = new THREE.Vector3(tx, ty, tz);
        const dummyCam = cameraOrtho.clone();
        dummyCam.position.copy(endPos);
        dummyCam.up.copy(targetUp);
        dummyCam.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        const endRot = dummyCam.quaternion.clone();
        const animData = { t: 0 };
        new TWEEN(animData)
            .to({ t: 1 }, 2000)
            .easing(t => t)
            .onUpdate(() => {
                const k = animData.t;
                cameraOrtho.position.lerpVectors(startPos, endPos, k);
                cameraOrtho.up.lerpVectors(startUp, targetUp, k).normalize();
                cameraOrtho.quaternion.slerpQuaternions(startRot, endRot, k);
            })
            .onComplete(() => {
                camSwitching = false;
                cameraOrtho.position.copy(endPos);
                cameraOrtho.up.copy(targetUp);
                cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.update();
                applyLockState();
            })
            .start();
    }
};
// Save Camera 3 state before mobile persp toggle overwrites the shared perspective camera
let savedCam3State = null;
window.toggleMobileCameraMode = function() {
    // Only allow toggle when on mobile camera
    if (currentCameraIndex !== 'mobile') return;

    if (currentCameraMode === 'ortho') {
        // Switch mobile to perspective - save Camera 3 state first
        savedCam3State = {
            position: camera.position.clone(),
            fov: camera.fov,
            target: controls.target.clone()
        };
        currentCameraMode = 'persp';
        if (transformControls) transformControls.camera = camera;
        // Transfer position from ortho camera to perspective camera
        camera.position.copy(cameraOrtho.position);
        camera.lookAt(controlsOrtho.target);
        controls.target.copy(controlsOrtho.target);
        camera.fov = runtimeParams.camMobileFov || 45;
        camera.updateProjectionMatrix();
        controls.update();
        applyLockState();
    } else {
        // Switch mobile back to orthographic - restore Camera 3 state
        currentCameraMode = 'ortho';
        if (transformControls) transformControls.camera = cameraOrtho;
        // Transfer position from perspective camera to ortho camera
        cameraOrtho.position.copy(camera.position);
        cameraOrtho.up.set(0, 1, 0);
        cameraOrtho.lookAt(controls.target);
        controlsOrtho.target.copy(controls.target);
        updateOrthoFrustum();
        cameraOrtho.updateProjectionMatrix();
        controlsOrtho.update();
        // Restore Camera 3's perspective camera to its original state
        if (savedCam3State) {
            camera.position.copy(savedCam3State.position);
            camera.fov = savedCam3State.fov;
            camera.updateProjectionMatrix();
            controls.target.copy(savedCam3State.target);
            controls.update();
            savedCam3State = null;
        }
        applyLockState();
    }
};
function getActiveCamera() {
    return currentCameraMode === 'persp' ? camera : cameraOrtho;
}
function handleEditorClick(e) {
    if (!editorModeActive) return;
    mouse.x = (e.clientX / window.innerWidth) * 2 - 1;
    mouse.y = -(e.clientY / window.innerHeight) * 2 + 1;
    const activeCamera = getActiveCamera();
    raycaster.setFromCamera(mouse, activeCamera);
    const hits = raycaster.intersectObjects(clickableObjects, true);
    if (hits.length > 0) {
        let obj = hits[0].object;
        while (obj.parent && !obj.userData.modalContent) obj = obj.parent;
        if (obj.userData.sceneObjectIndex !== undefined) selectEditorObject(obj);
    }
}
window.printEditorData = function() {
    let output = ""; 
    sceneObjects.forEach((item, index) => {
        const obj = loadedObjectsMap.get(index);
        if (!obj) return;
        const pos = obj.position;
        const rot = obj.rotation;
        const scl = obj.scale.x;
        output += `${pos.x.toFixed(0)}\t${pos.y.toFixed(0)}\t${pos.z.toFixed(0)}\t`;
        output += `${(rot.x * 180 / Math.PI).toFixed(0)}\t${(rot.y * 180 / Math.PI).toFixed(0)}\t${(rot.z * 180 / Math.PI).toFixed(0)}\t`;
        output += `${scl.toFixed(0)}\n`;
    });
    navigator.clipboard.writeText(output).then(() => {
        alert('COPIED TO CLIPBOARD!\n\nPaste this directly into your Excel sheet starting at cell AB8 (Columns AB -> AH).');
    }).catch(err => {
        console.log(output);
        alert('Copy failed. Check console for data.');
    });
};
window.switchEditorTab = function(tabName) {
    document.querySelectorAll('.editor-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.tab === tabName);
    });
    document.querySelectorAll('.editor-tab-content').forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabName}`);
    });
    if (tabName === 'parameters') {
        buildParametersPanel();
    } else if (tabName === 'scene') {
        buildSceneParametersPanel();
    }
};
function attachValueEditing(container) {
    container.querySelectorAll('.params-slider-value').forEach(span => {
        if (span.dataset.editBound) return;
        span.dataset.editBound = '1';
        span.style.cursor = 'pointer';
        span.addEventListener('click', function() {
            const row = this.closest('.params-slider-row');
            if (!row) return;
            const slider = row.querySelector('.params-slider');
            if (!slider) return;
            const input = document.createElement('input');
            input.type = 'number';
            input.className = 'params-slider-input';
            input.value = parseFloat(slider.value);
            input.step = 'any';
            const origSpan = this;
            const parentRow = this.parentElement;
            const finishEdit = () => {
                const newValue = parseFloat(input.value);
                if (!isNaN(newValue)) {
                    slider.value = newValue;
                    slider.dispatchEvent(new Event('change'));
                }
                if (input.parentElement) {
                    const newSpan = origSpan.cloneNode(false);
                    newSpan.textContent = isNaN(newValue) ? origSpan.textContent : newValue;
                    newSpan.dataset.editBound = '';
                    input.replaceWith(newSpan);
                    attachValueEditing(parentRow);
                }
            };
            input.onblur = finishEdit;
            input.onkeydown = (e) => {
                e.stopPropagation();
                if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); finishEdit(); }
                if (e.key === 'Escape') {
                    if (input.parentElement) input.replaceWith(origSpan);
                    origSpan.dataset.editBound = '';
                    attachValueEditing(parentRow);
                }
            };
            this.replaceWith(input);
            input.focus();
            input.select();
        });
    });
}
function buildParametersPanel() {
    const grid = document.getElementById('params-grid');
    const isMobile = isMobileView();
    grid.innerHTML = `
        <div class="params-section-title">SSAO (Ambient Occlusion)</div>
        <span class="params-label">SSAO Desktop</span>
        <span class="params-value ${runtimeParams.ssaoEnabledDesktop ? 'on' : 'off'}">${runtimeParams.ssaoEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.ssaoEnabledDesktop ? 'on' : ''}" onclick="toggleParam('ssaoEnabledDesktop')"></div>
        <span class="params-label">SSAO Mobile</span>
        <span class="params-value ${runtimeParams.ssaoEnabledMobile ? 'on' : 'off'}">${runtimeParams.ssaoEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.ssaoEnabledMobile ? 'on' : ''}" onclick="toggleParam('ssaoEnabledMobile')"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Kernel Radius</span>
            <input type="range" class="params-slider" min="1" max="128" step="1" value="${runtimeParams.ssaoKernelRadius}" onchange="updateSliderParam('ssaoKernelRadius', this.value)">
            <span class="params-slider-value">${runtimeParams.ssaoKernelRadius}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Intensity</span>
            <input type="range" class="params-slider" min="0" max="5" step="0.1" value="${runtimeParams.ssaoIntensity}" onchange="updateSliderParam('ssaoIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.ssaoIntensity}</span>
        </div>
        <div class="params-section-title">BLOOM (Glow)</div>
        <span class="params-label">Bloom Desktop</span>
        <span class="params-value ${runtimeParams.bloomEnabledDesktop ? 'on' : 'off'}">${runtimeParams.bloomEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.bloomEnabledDesktop ? 'on' : ''}" onclick="toggleParam('bloomEnabledDesktop')"></div>
        <span class="params-label">Bloom Mobile</span>
        <span class="params-value ${runtimeParams.bloomEnabledMobile ? 'on' : 'off'}">${runtimeParams.bloomEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.bloomEnabledMobile ? 'on' : ''}" onclick="toggleParam('bloomEnabledMobile')"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Strength</span>
            <input type="range" class="params-slider" min="0" max="0.5" step="0.01" value="${runtimeParams.bloomStrength}" onchange="updateSliderParam('bloomStrength', this.value)">
            <span class="params-slider-value">${runtimeParams.bloomStrength}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Threshold</span>
            <input type="range" class="params-slider" min="0" max="2" step="0.01" value="${runtimeParams.bloomThreshold}" onchange="updateSliderParam('bloomThreshold', this.value)">
            <span class="params-slider-value">${runtimeParams.bloomThreshold}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Radius</span>
            <input type="range" class="params-slider" min="0" max="2" step="0.01" value="${runtimeParams.bloomRadius}" onchange="updateSliderParam('bloomRadius', this.value)">
            <span class="params-slider-value">${runtimeParams.bloomRadius}</span>
        </div>
        <div class="params-section-title">TONE MAPPING</div>
        <span class="params-label">Tone Mapping Desktop</span>
        <span class="params-value ${runtimeParams.tonemappingEnabledDesktop ? 'on' : 'off'}">${runtimeParams.tonemappingEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.tonemappingEnabledDesktop ? 'on' : ''}" onclick="toggleParam('tonemappingEnabledDesktop')"></div>
        <span class="params-label">Tone Mapping Mobile</span>
        <span class="params-value ${runtimeParams.tonemappingEnabledMobile ? 'on' : 'off'}">${runtimeParams.tonemappingEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.tonemappingEnabledMobile ? 'on' : ''}" onclick="toggleParam('tonemappingEnabledMobile')"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Exposure Desktop</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.05" value="${runtimeParams.tonemappingExposureDesktop}" onchange="updateSliderParam('tonemappingExposureDesktop', this.value)">
            <span class="params-slider-value">${runtimeParams.tonemappingExposureDesktop}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Exposure Mobile</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.05" value="${runtimeParams.tonemappingExposureMobile}" onchange="updateSliderParam('tonemappingExposureMobile', this.value)">
            <span class="params-slider-value">${runtimeParams.tonemappingExposureMobile}</span>
        </div>
        <div class="params-section-title">SHADOWS</div>
        <span class="params-label">Shadows Desktop</span>
        <span class="params-value ${runtimeParams.shadowsEnabledDesktop ? 'on' : 'off'}">${runtimeParams.shadowsEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.shadowsEnabledDesktop ? 'on' : ''}" onclick="toggleParam('shadowsEnabledDesktop')"></div>
        <span class="params-label">Shadows Mobile</span>
        <span class="params-value ${runtimeParams.shadowsEnabledMobile ? 'on' : 'off'}">${runtimeParams.shadowsEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.shadowsEnabledMobile ? 'on' : ''}" onclick="toggleParam('shadowsEnabledMobile')"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Shadow Bias</span>
            <input type="range" class="params-slider" min="-0.01" max="0.01" step="0.0001" value="${runtimeParams.shadowBias}" onchange="updateSliderParam('shadowBias', this.value)">
            <span class="params-slider-value">${runtimeParams.shadowBias}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Normal Bias</span>
            <input type="range" class="params-slider" min="0" max="1" step="0.01" value="${runtimeParams.shadowNormalBias}" onchange="updateSliderParam('shadowNormalBias', this.value)">
            <span class="params-slider-value">${runtimeParams.shadowNormalBias}</span>
        </div>
        <div class="params-section-title">ANTI-ALIASING</div>
        <span class="params-label">FXAA Desktop</span>
        <span class="params-value ${runtimeParams.fxaaEnabledDesktop ? 'on' : 'off'}">${runtimeParams.fxaaEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.fxaaEnabledDesktop ? 'on' : ''}" onclick="toggleParam('fxaaEnabledDesktop')"></div>
        <span class="params-label">FXAA Mobile</span>
        <span class="params-value ${runtimeParams.fxaaEnabledMobile ? 'on' : 'off'}">${runtimeParams.fxaaEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.fxaaEnabledMobile ? 'on' : ''}" onclick="toggleParam('fxaaEnabledMobile')"></div>
        <div class="params-section-title">ANISOTROPY (Texture Quality)</div>
        <span class="params-label">Anisotropy Desktop</span>
        <span class="params-value ${runtimeParams.anisotropyEnabledDesktop ? 'on' : 'off'}">${runtimeParams.anisotropyEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.anisotropyEnabledDesktop ? 'on' : ''}" onclick="toggleParam('anisotropyEnabledDesktop')"></div>
        <span class="params-label">Anisotropy Mobile</span>
        <span class="params-value ${runtimeParams.anisotropyEnabledMobile ? 'on' : 'off'}">${runtimeParams.anisotropyEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.anisotropyEnabledMobile ? 'on' : ''}" onclick="toggleParam('anisotropyEnabledMobile')"></div>
        <div class="params-section-title">ENVIRONMENT MAP</div>
        <span class="params-label">EnvMap Desktop</span>
        <span class="params-value ${runtimeParams.envmapEnabledDesktop ? 'on' : 'off'}">${runtimeParams.envmapEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.envmapEnabledDesktop ? 'on' : ''}" onclick="toggleParam('envmapEnabledDesktop')"></div>
        <span class="params-label">EnvMap Mobile</span>
        <span class="params-value ${runtimeParams.envmapEnabledMobile ? 'on' : 'off'}">${runtimeParams.envmapEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.envmapEnabledMobile ? 'on' : ''}" onclick="toggleParam('envmapEnabledMobile')"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Intensity Desktop</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.1" value="${runtimeParams.envmapIntensityDesktop}" onchange="updateSliderParam('envmapIntensityDesktop', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'envmapIntensityDesktop')">${runtimeParams.envmapIntensityDesktop}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Intensity Mobile</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.1" value="${runtimeParams.envmapIntensityMobile}" onchange="updateSliderParam('envmapIntensityMobile', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'envmapIntensityMobile')">${runtimeParams.envmapIntensityMobile}</span>
        </div>
        <div class="params-section-title">RECT AREA LIGHT</div>
        <span class="params-label">RectAreaLight Desktop</span>
        <span class="params-value ${runtimeParams.rectarealightEnabledDesktop ? 'on' : 'off'}">${runtimeParams.rectarealightEnabledDesktop ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.rectarealightEnabledDesktop ? 'on' : ''}" onclick="toggleParam('rectarealightEnabledDesktop')"></div>
        <span class="params-label">RectAreaLight Mobile</span>
        <span class="params-value ${runtimeParams.rectarealightEnabledMobile ? 'on' : 'off'}">${runtimeParams.rectarealightEnabledMobile ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.rectarealightEnabledMobile ? 'on' : ''}" onclick="toggleParam('rectarealightEnabledMobile')"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Intensity</span>
            <input type="range" class="params-slider" min="0" max="20" step="0.1" value="${runtimeParams.rectarealightIntensity}" onchange="updateSliderParam('rectarealightIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.rectarealightIntensity}</span>
        </div>
        <div class="params-section-title">LIGHTING</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Hemisphere</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.05" value="${runtimeParams.lightHemisphereIntensity}" onchange="updateSliderParam('lightHemisphereIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.lightHemisphereIntensity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Sun</span>
            <input type="range" class="params-slider" min="0" max="5" step="0.1" value="${runtimeParams.lightSunIntensity}" onchange="updateSliderParam('lightSunIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.lightSunIntensity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Sun X</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightSunPosX}" onchange="updateSliderParam('lightSunPosX', this.value)">
            <span class="params-slider-value">${runtimeParams.lightSunPosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Sun Y</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightSunPosY}" onchange="updateSliderParam('lightSunPosY', this.value)">
            <span class="params-slider-value">${runtimeParams.lightSunPosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Sun Z</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightSunPosZ}" onchange="updateSliderParam('lightSunPosZ', this.value)">
            <span class="params-slider-value">${runtimeParams.lightSunPosZ}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Fill</span>
            <input type="range" class="params-slider" min="0" max="2" step="0.1" value="${runtimeParams.lightFillIntensity}" onchange="updateSliderParam('lightFillIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.lightFillIntensity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Fill X</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightFillPosX}" onchange="updateSliderParam('lightFillPosX', this.value)">
            <span class="params-slider-value">${runtimeParams.lightFillPosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Fill Y</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightFillPosY}" onchange="updateSliderParam('lightFillPosY', this.value)">
            <span class="params-slider-value">${runtimeParams.lightFillPosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Fill Z</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightFillPosZ}" onchange="updateSliderParam('lightFillPosZ', this.value)">
            <span class="params-slider-value">${runtimeParams.lightFillPosZ}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rim</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.05" value="${runtimeParams.lightRimIntensity}" onchange="updateSliderParam('lightRimIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.lightRimIntensity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rim X</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightRimPosX}" onchange="updateSliderParam('lightRimPosX', this.value)">
            <span class="params-slider-value">${runtimeParams.lightRimPosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rim Y</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightRimPosY}" onchange="updateSliderParam('lightRimPosY', this.value)">
            <span class="params-slider-value">${runtimeParams.lightRimPosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rim Z</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightRimPosZ}" onchange="updateSliderParam('lightRimPosZ', this.value)">
            <span class="params-slider-value">${runtimeParams.lightRimPosZ}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Accent</span>
            <input type="range" class="params-slider" min="0" max="10" step="0.05" value="${runtimeParams.lightAccentIntensity}" onchange="updateSliderParam('lightAccentIntensity', this.value)">
            <span class="params-slider-value">${runtimeParams.lightAccentIntensity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Accent X</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightAccentPosX}" onchange="updateSliderParam('lightAccentPosX', this.value)">
            <span class="params-slider-value">${runtimeParams.lightAccentPosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Accent Y</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightAccentPosY}" onchange="updateSliderParam('lightAccentPosY', this.value)">
            <span class="params-slider-value">${runtimeParams.lightAccentPosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Accent Z</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="100" value="${runtimeParams.lightAccentPosZ}" onchange="updateSliderParam('lightAccentPosZ', this.value)">
            <span class="params-slider-value">${runtimeParams.lightAccentPosZ}</span>
        </div>
        <div class="params-section-title">PBR MATERIALS</div>
        <div class="params-slider-row">
            <span class="params-slider-label">EnvMap Intensity</span>
            <input type="range" class="params-slider" min="0" max="5" step="0.1" value="${runtimeParams.pbrEnvmapIntensity}" onchange="updateSliderParam('pbrEnvmapIntensity', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'pbrEnvmapIntensity')">${runtimeParams.pbrEnvmapIntensity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Roughness Boost</span>
            <input type="range" class="params-slider" min="-0.5" max="0.5" step="0.05" value="${runtimeParams.pbrRoughnessBoost}" onchange="updateSliderParam('pbrRoughnessBoost', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'pbrRoughnessBoost')">${runtimeParams.pbrRoughnessBoost}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Metalness Boost</span>
            <input type="range" class="params-slider" min="-0.5" max="0.5" step="0.05" value="${runtimeParams.pbrMetalnessBoost}" onchange="updateSliderParam('pbrMetalnessBoost', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'pbrMetalnessBoost')">${runtimeParams.pbrMetalnessBoost}</span>
        </div>
        <div class="params-section-title">COLOR CORRECTION</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Brightness</span>
            <input type="range" class="params-slider" min="0.5" max="2" step="0.05" value="${runtimeParams.filterBrightness}" onchange="updateSliderParam('filterBrightness', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'filterBrightness')">${runtimeParams.filterBrightness}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Contrast</span>
            <input type="range" class="params-slider" min="0" max="5" step="0.05" value="${runtimeParams.filterContrast}" onchange="updateSliderParam('filterContrast', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'filterContrast')">${runtimeParams.filterContrast}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Saturation</span>
            <input type="range" class="params-slider" min="0" max="5" step="0.05" value="${runtimeParams.filterSaturation}" onchange="updateSliderParam('filterSaturation', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'filterSaturation')">${runtimeParams.filterSaturation}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Hue Rotate</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${runtimeParams.filterHueRotate}" onchange="updateSliderParam('filterHueRotate', this.value)" oninput="updateSliderDisplay(this)">
            <span class="params-slider-value" onclick="makeEditable(this, 'filterHueRotate')">${runtimeParams.filterHueRotate}°</span>
        </div>
        <div class="params-section-title">HDRI PRESET (Shift+1..6)</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Desktop: ${currentHdriPreset} - ${hdriPresets[currentHdriPreset]?.name || 'Unknown'}</span>
            <input type="range" class="params-slider" min="1" max="6" step="1" value="${currentHdriPreset}" onchange="switchHdriPreset(parseInt(this.value), false)">
            <span class="params-slider-value">${currentHdriPreset}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Mobile: ${currentHdriPresetMobile} - ${hdriPresets[currentHdriPresetMobile]?.name || 'Unknown'}</span>
            <input type="range" class="params-slider" min="1" max="6" step="1" value="${currentHdriPresetMobile}" onchange="switchHdriPreset(parseInt(this.value), true)">
            <span class="params-slider-value">${currentHdriPresetMobile}</span>
        </div>
    `;
    attachValueEditing(grid);
}
window.buildSceneParametersPanel = function() {
    const grid = document.getElementById('scene-params-grid');
    const envObj = sceneObjects.find(o => o.modalContent && o.modalContent.name === 'Environment');
    let envBarHtml = '';
    if (envObj && envObj.envOptions) {
        const opts = envObj.envOptions;
        let btns = '';
        let stopped = false;
        for (let i = 0; i < 16; i++) {
            if (stopped || opts[i] === null || opts[i] === undefined) { stopped = true; continue; }
            const isActive = (runtimeParams.envOption || 1) === (i + 1);
            btns += `<button class="env-option-btn${isActive ? ' active' : ''}" data-env-idx="${i + 1}" onclick="selectEnvOption(${i + 1})">${i + 1}</button>`;
        }
        if (btns) {
            envBarHtml = `<div class="env-option-bar"><span class="env-option-label">Environment Option</span>${btns}</div>`;
        }
    }
    grid.innerHTML = envBarHtml + `
        <div class="params-section-title">CAMERA 1 POSITION</div>
        <div class="params-slider-row">
            <span class="params-slider-label">X</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam1PosX}" onchange="updateSceneSliderParam('cam1PosX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam1PosX')">${runtimeParams.cam1PosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Y</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam1PosY}" onchange="updateSceneSliderParam('cam1PosY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam1PosY')">${runtimeParams.cam1PosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Z</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam1PosZ}" onchange="updateSceneSliderParam('cam1PosZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam1PosZ')">${runtimeParams.cam1PosZ}</span>
        </div>
        <span class="params-label">Cam 1 Lock Mode</span>
        <select onchange="runtimeParams.cam1LockMode = this.value; applyLockState(); buildSceneParametersPanel()" style="background:#111;color:#fff;border:1px solid #333;padding:2px 4px;font-size:11px;font-family:monospace;cursor:pointer;">
            <option value="none" ${runtimeParams.cam1LockMode === 'none' ? 'selected' : ''}>none</option>
            <option value="horizontal" ${runtimeParams.cam1LockMode === 'horizontal' ? 'selected' : ''}>horizontal</option>
            <option value="full" ${runtimeParams.cam1LockMode === 'full' ? 'selected' : ''}>full</option>
        </select>
        <span class="params-slider-value">${runtimeParams.cam1LockMode}</span>
        <span class="params-label">Show Cam 2 Public</span>
        <span class="params-value ${runtimeParams.cam2ShowPublic ? 'on' : 'off'}">${runtimeParams.cam2ShowPublic ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.cam2ShowPublic ? 'on' : ''}" onclick="toggleParam('cam2ShowPublic'); updateCameraDotsVisibility(); buildSceneParametersPanel()"></div>
        <div class="params-section-title">CAMERA 2 POSITION</div>
        <div class="params-slider-row">
            <span class="params-slider-label">X</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam2PosX}" onchange="updateSceneSliderParam('cam2PosX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam2PosX')">${runtimeParams.cam2PosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Y</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam2PosY}" onchange="updateSceneSliderParam('cam2PosY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam2PosY')">${runtimeParams.cam2PosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Z</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam2PosZ}" onchange="updateSceneSliderParam('cam2PosZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam2PosZ')">${runtimeParams.cam2PosZ}</span>
        </div>
        <span class="params-label">Cam 2 Lock Mode</span>
        <select onchange="runtimeParams.cam2LockMode = this.value; applyLockState(); buildSceneParametersPanel()" style="background:#111;color:#fff;border:1px solid #333;padding:2px 4px;font-size:11px;font-family:monospace;cursor:pointer;">
            <option value="none" ${runtimeParams.cam2LockMode === 'none' ? 'selected' : ''}>none</option>
            <option value="horizontal" ${runtimeParams.cam2LockMode === 'horizontal' ? 'selected' : ''}>horizontal</option>
            <option value="full" ${runtimeParams.cam2LockMode === 'full' ? 'selected' : ''}>full</option>
        </select>
        <span class="params-slider-value">${runtimeParams.cam2LockMode}</span>
        <div class="params-section-title">CAMERA 3 POSITION</div>
        <div class="params-slider-row">
            <span class="params-slider-label">X</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam3PosX}" onchange="updateSceneSliderParam('cam3PosX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam3PosX')">${runtimeParams.cam3PosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Y</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam3PosY}" onchange="updateSceneSliderParam('cam3PosY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam3PosY')">${runtimeParams.cam3PosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Z</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.cam3PosZ}" onchange="updateSceneSliderParam('cam3PosZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cam3PosZ')">${runtimeParams.cam3PosZ}</span>
        </div>
        <span class="params-label">Cam 3 Lock Mode</span>
        <select onchange="runtimeParams.cam3LockMode = this.value; applyLockState(); buildSceneParametersPanel()" style="background:#111;color:#fff;border:1px solid #333;padding:2px 4px;font-size:11px;font-family:monospace;cursor:pointer;">
            <option value="none" ${runtimeParams.cam3LockMode === 'none' ? 'selected' : ''}>none</option>
            <option value="horizontal" ${runtimeParams.cam3LockMode === 'horizontal' ? 'selected' : ''}>horizontal</option>
            <option value="full" ${runtimeParams.cam3LockMode === 'full' ? 'selected' : ''}>full</option>
        </select>
        <span class="params-slider-value">${runtimeParams.cam3LockMode}</span>
        <span class="params-label">Show Cam 3 Public</span>
        <span class="params-value ${runtimeParams.cam3ShowPublic ? 'on' : 'off'}">${runtimeParams.cam3ShowPublic ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.cam3ShowPublic ? 'on' : ''}" onclick="toggleParam('cam3ShowPublic'); updateCam3Visibility(); updateCameraDotsVisibility(); buildSceneParametersPanel()"></div>
        <span class="params-label">Mobile Labels Public</span>
        <span class="params-value ${runtimeParams.mobileLabelsShowPublic ? 'on' : 'off'}">${runtimeParams.mobileLabelsShowPublic ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.mobileLabelsShowPublic ? 'on' : ''}" onclick="toggleParam('mobileLabelsShowPublic'); buildSceneParametersPanel()"></div>
        <div class="params-section-title">MOBILE CAMERA</div>
        <span class="params-label">Camera Mode</span>
        <span class="params-value">${currentCameraMode === 'persp' ? 'PERSPECTIVE' : 'ORTHOGRAPHIC'}</span>
        <div class="params-toggle ${currentCameraMode === 'persp' ? 'on' : ''}" onclick="toggleMobileCameraMode(); buildSceneParametersPanel()"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">X</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.camMobilePosX}" onchange="updateSceneSliderParam('camMobilePosX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camMobilePosX')">${runtimeParams.camMobilePosX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Y</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.camMobilePosY}" onchange="updateSceneSliderParam('camMobilePosY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camMobilePosY')">${runtimeParams.camMobilePosY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Z</span>
            <input type="range" class="params-slider" min="-50000" max="50000" step="100" value="${runtimeParams.camMobilePosZ}" onchange="updateSceneSliderParam('camMobilePosZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camMobilePosZ')">${runtimeParams.camMobilePosZ}</span>
        </div>
        <span class="params-label">Mobile Cam Lock Mode</span>
        <select onchange="runtimeParams.camMobileLockMode = this.value; applyLockState(); buildSceneParametersPanel()" style="background:#111;color:#fff;border:1px solid #333;padding:2px 4px;font-size:11px;font-family:monospace;cursor:pointer;">
            <option value="none" ${runtimeParams.camMobileLockMode === 'none' ? 'selected' : ''}>none</option>
            <option value="horizontal" ${runtimeParams.camMobileLockMode === 'horizontal' ? 'selected' : ''}>horizontal</option>
            <option value="full" ${runtimeParams.camMobileLockMode === 'full' ? 'selected' : ''}>full</option>
        </select>
        <span class="params-slider-value">${runtimeParams.camMobileLockMode}</span>
        <div class="params-slider-row">
            <span class="params-slider-label">Mobile FOV</span>
            <input type="range" class="params-slider" min="10" max="120" step="1" value="${runtimeParams.camMobileFov}" onchange="updateSceneSliderParam('camMobileFov', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camMobileFov')">${runtimeParams.camMobileFov}</span>
        </div>
        <div class="params-section-title">CAMERA SETTINGS</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Initial Zoom</span>
            <input type="range" class="params-slider" min="0.5" max="10" step="0.1" value="${runtimeParams.initialCameraZoom}" onchange="updateSceneSliderParam('initialCameraZoom', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'initialCameraZoom')">${runtimeParams.initialCameraZoom}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Initial Zoom Mobile</span>
            <input type="range" class="params-slider" min="0.5" max="10" step="0.1" value="${runtimeParams.initialCameraZoomMobile}" onchange="updateSceneSliderParam('initialCameraZoomMobile', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'initialCameraZoomMobile')">${runtimeParams.initialCameraZoomMobile}</span>
        </div>
        <span class="params-label">Lock Button Public</span>
        <span class="params-value ${runtimeParams.cameraLockShowPublic ? 'on' : 'off'}">${runtimeParams.cameraLockShowPublic ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.cameraLockShowPublic ? 'on' : ''}" onclick="runtimeParams.cameraLockShowPublic = !runtimeParams.cameraLockShowPublic; const lb = document.getElementById('btn-lock-icon-public'); if(lb) lb.style.display = runtimeParams.cameraLockShowPublic ? '' : 'none'; buildSceneParametersPanel()"></div>
        <span class="params-label">Default Locked</span>
        <span class="params-value ${runtimeParams.cameraDefaultLocked ? 'on' : 'off'}">${runtimeParams.cameraDefaultLocked ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.cameraDefaultLocked ? 'on' : ''}" onclick="runtimeParams.cameraDefaultLocked = !runtimeParams.cameraDefaultLocked; buildSceneParametersPanel()"></div>
        <span class="params-label">Zoom Rotate</span>
        <span class="params-value ${runtimeParams.zoomRotateEnabled ? 'on' : 'off'}">${runtimeParams.zoomRotateEnabled ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${runtimeParams.zoomRotateEnabled ? 'on' : ''}" onclick="runtimeParams.zoomRotateEnabled = !runtimeParams.zoomRotateEnabled; buildSceneParametersPanel()"></div>
        <div class="params-section-title">PANEL SETTINGS</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Panel Width</span>
            <input type="range" class="params-slider" min="200" max="500" step="10" value="${runtimeParams.panelWidth}" onchange="updateSceneSliderParam('panelWidth', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'panelWidth')">${runtimeParams.panelWidth}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Panel Top Offset</span>
            <input type="range" class="params-slider" min="0" max="200" step="5" value="${runtimeParams.panelTopOffset}" onchange="updateSceneSliderParam('panelTopOffset', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'panelTopOffset')">${runtimeParams.panelTopOffset}px</span>
        </div>
        <div class="params-section-title">MARQUEE TEXT</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Speed (s)</span>
            <input type="range" class="params-slider" min="5" max="60" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-speed')) || 20}" onchange="updateMarqueeParam('speed', this.value, this)">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-speed')) || 20}s</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Font Size (px)</span>
            <input type="range" class="params-slider" min="6" max="24" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-font-size')) || 9}" onchange="updateMarqueeParam('fontSize', this.value, this)">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-font-size')) || 9}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Vertical Gap (px)</span>
            <input type="range" class="params-slider" min="0" max="30" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-vertical-gap')) || 6}" onchange="updateMarqueeParam('verticalGap', this.value, this)">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-vertical-gap')) || 6}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Color</span>
            <input type="color" value="${getComputedStyle(document.documentElement).getPropertyValue('--marquee-color').trim() || '#b6a6d0'}" onchange="updateMarqueeParam('color', this.value, this)" style="width:32px;height:20px;border:none;cursor:pointer;padding:0;">
            <span class="params-slider-value">${getComputedStyle(document.documentElement).getPropertyValue('--marquee-color').trim() || '#b6a6d0'}</span>
        </div>
        <div class="params-section-title">MARQUEE MOBILE</div>
        <span class="params-label">Mobile Marquee</span>
        <span class="params-value ${mobileMarqueeEnabled ? 'on' : 'off'}">${mobileMarqueeEnabled ? 'ON' : 'OFF'}</span>
        <div class="params-toggle ${mobileMarqueeEnabled ? 'on' : ''}" onclick="mobileMarqueeEnabled = !mobileMarqueeEnabled; positionMarquee(); buildSceneParametersPanel()"></div>
        <div class="params-slider-row">
            <span class="params-slider-label">Mobile Speed (s)</span>
            <input type="range" class="params-slider" min="5" max="60" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-speed')) || 20}" onchange="updateMarqueeParam('mobileSpeed', this.value, this)">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-speed')) || 20}s</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Mobile Font (px)</span>
            <input type="range" class="params-slider" min="4" max="18" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-font-size')) || 7}" onchange="updateMarqueeParam('mobileFontSize', this.value, this)">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-font-size')) || 7}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Mobile Top (px)</span>
            <input type="range" class="params-slider" min="0" max="50" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-top-offset')) || 12}" onchange="updateMarqueeParam('mobileTopOffset', this.value, this)">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--marquee-mobile-top-offset')) || 12}px</span>
        </div>
        <div class="params-section-title">CAMERA LIMITS</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Min Polar Angle (°)</span>
            <input type="range" class="params-slider" min="0" max="90" step="1" value="${runtimeParams.cameraMinPolarAngle}" onchange="updateSceneSliderParam('cameraMinPolarAngle', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cameraMinPolarAngle')">${runtimeParams.cameraMinPolarAngle}°</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Max Polar Angle (°)</span>
            <input type="range" class="params-slider" min="0" max="90" step="1" value="${runtimeParams.cameraMaxPolarAngle}" onchange="updateSceneSliderParam('cameraMaxPolarAngle', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cameraMaxPolarAngle')">${runtimeParams.cameraMaxPolarAngle}°</span>
        </div>
        <div class="params-section-title">CAMERA GENERAL</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Perspective FOV</span>
            <input type="range" class="params-slider" min="10" max="120" step="1" value="${PARAMS.cameraPerspectiveFov}" onchange="updateSceneSliderParam('cameraPerspectiveFov', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cameraPerspectiveFov')">${PARAMS.cameraPerspectiveFov}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Ortho FOV</span>
            <input type="range" class="params-slider" min="10" max="120" step="1" value="${PARAMS.cameraOrthoFov}" onchange="updateSceneSliderParam('cameraOrthoFov', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cameraOrthoFov')">${PARAMS.cameraOrthoFov}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Min Distance</span>
            <input type="range" class="params-slider" min="1" max="5000" step="10" value="${PARAMS.cameraMinDist}" onchange="updateSceneSliderParam('cameraMinDist', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cameraMinDist')">${PARAMS.cameraMinDist}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Max Distance</span>
            <input type="range" class="params-slider" min="1000" max="100000" step="500" value="${PARAMS.cameraMaxDist}" onchange="updateSceneSliderParam('cameraMaxDist', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'cameraMaxDist')">${PARAMS.cameraMaxDist}</span>
        </div>
        <span class="params-label">Enable Pan</span>
        <span class="params-value ${PARAMS.enablePan ? 'on' : 'off'}">${PARAMS.enablePan ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${PARAMS.enablePan ? 'on' : ''}" onclick="PARAMS.enablePan = !PARAMS.enablePan; if(controls) controls.enablePan = PARAMS.enablePan; if(controlsOrtho) controlsOrtho.enablePan = PARAMS.enablePan; buildSceneParametersPanel()"></div>
        <span class="params-label">Mobile Pan</span>
        <span class="params-value ${PARAMS.mobilePanEnabled ? 'on' : 'off'}">${PARAMS.mobilePanEnabled ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${PARAMS.mobilePanEnabled ? 'on' : ''}" onclick="PARAMS.mobilePanEnabled = !PARAMS.mobilePanEnabled; const mp = PARAMS.mobilePanEnabled; if(controls) controls.enablePan = mp; if(controlsOrtho) controlsOrtho.enablePan = mp; if(mp) { controls.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_PAN }; controlsOrtho.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_PAN }; } else { controls.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_ROTATE }; controlsOrtho.touches = { ONE: THREE.TOUCH.ROTATE, TWO: THREE.TOUCH.DOLLY_ROTATE }; } buildSceneParametersPanel()"></div>
        <div class="params-section-title">CAM LOOK-AT</div>
        <div class="params-slider-row">
            <span class="params-slider-label">X</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="10" value="${PARAMS.camLookX}" onchange="updateSceneSliderParam('camLookX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camLookX')">${PARAMS.camLookX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Y</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="10" value="${PARAMS.camLookY}" onchange="updateSceneSliderParam('camLookY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camLookY')">${PARAMS.camLookY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Z</span>
            <input type="range" class="params-slider" min="-5000" max="5000" step="10" value="${PARAMS.camLookZ}" onchange="updateSceneSliderParam('camLookZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'camLookZ')">${PARAMS.camLookZ}</span>
        </div>
        <div class="params-section-title">HOVER ANIMATION</div>
        <div class="params-slider-row">
            <span class="params-slider-label">While-Selected Factor</span>
            <input type="range" class="params-slider" min="0.05" max="1" step="0.05" value="${PARAMS.hoverWhileSelectedFactor}" onchange="updateSceneSliderParam('hoverWhileSelectedFactor', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverWhileSelectedFactor')">${PARAMS.hoverWhileSelectedFactor}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Lift Distance</span>
            <input type="range" class="params-slider" min="0" max="500" step="5" value="${PARAMS.hoverLiftDistance}" onchange="updateSceneSliderParam('hoverLiftDistance', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLiftDistance')">${PARAMS.hoverLiftDistance}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Anim Duration (ms)</span>
            <input type="range" class="params-slider" min="100" max="2000" step="50" value="${PARAMS.hoverAnimationDuration}" onchange="updateSceneSliderParam('hoverAnimationDuration', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverAnimationDuration')">${PARAMS.hoverAnimationDuration}ms</span>
        </div>
        <div class="params-section-title">HOVER LABEL STYLE</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Text Size</span>
            <input type="range" class="params-slider" min="8" max="120" step="1" value="${PARAMS.hoverLabelFontSize}" onchange="updateSceneSliderParam('hoverLabelFontSize', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelFontSize')">${PARAMS.hoverLabelFontSize}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Font Style</span>
            <select onchange="PARAMS.hoverLabelFontStyle = this.value; buildSceneParametersPanel()" style="background:#111;color:#fff;border:1px solid #333;padding:2px 4px;font-size:11px;font-family:monospace;cursor:pointer;">
                <option value="normal" ${PARAMS.hoverLabelFontStyle === 'normal' ? 'selected' : ''}>normal</option>
                <option value="bold" ${PARAMS.hoverLabelFontStyle === 'bold' ? 'selected' : ''}>bold</option>
                <option value="italic" ${PARAMS.hoverLabelFontStyle === 'italic' ? 'selected' : ''}>italic</option>
            </select>
            <span class="params-slider-value">${PARAMS.hoverLabelFontStyle}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Letter Spacing</span>
            <input type="range" class="params-slider" min="-5" max="20" step="0.5" value="${PARAMS.hoverLabelLetterSpacing}" onchange="updateSceneSliderParam('hoverLabelLetterSpacing', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelLetterSpacing')">${PARAMS.hoverLabelLetterSpacing}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Label Color</span>
            <input type="color" value="${PARAMS.hoverLabelColor}" onchange="PARAMS.hoverLabelColor = this.value; buildSceneParametersPanel()" style="width:32px;height:20px;border:none;cursor:pointer;padding:0;">
            <span class="params-slider-value">${PARAMS.hoverLabelColor}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Label Opacity</span>
            <input type="range" class="params-slider" min="0" max="1" step="0.05" value="${PARAMS.hoverLabelOpacity}" onchange="updateSceneSliderParam('hoverLabelOpacity', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelOpacity')">${PARAMS.hoverLabelOpacity}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Font Box</span>
            <input type="range" class="params-slider" min="100" max="5000" step="50" value="${PARAMS.hoverLabelFontBox}" onchange="updateSceneSliderParam('hoverLabelFontBox', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelFontBox')">${PARAMS.hoverLabelFontBox}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Model Gap (Offset Y)</span>
            <input type="range" class="params-slider" min="-500" max="500" step="5" value="${PARAMS.hoverLabelOffsetY}" onchange="updateSceneSliderParam('hoverLabelOffsetY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelOffsetY')">${PARAMS.hoverLabelOffsetY}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Artist Gap (Name Offset Y)</span>
            <input type="range" class="params-slider" min="-500" max="500" step="5" value="${PARAMS.hoverLabelNameOffsetY}" onchange="updateSceneSliderParam('hoverLabelNameOffsetY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelNameOffsetY')">${PARAMS.hoverLabelNameOffsetY}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotation X (°)</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${PARAMS.hoverLabelRotationX}" onchange="updateSceneSliderParam('hoverLabelRotationX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelRotationX')">${PARAMS.hoverLabelRotationX}°</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotation Y (°)</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${PARAMS.hoverLabelRotationY}" onchange="updateSceneSliderParam('hoverLabelRotationY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelRotationY')">${PARAMS.hoverLabelRotationY}°</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotation Z (°)</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${PARAMS.hoverLabelRotationZ}" onchange="updateSceneSliderParam('hoverLabelRotationZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverLabelRotationZ')">${PARAMS.hoverLabelRotationZ}°</span>
        </div>
        <div class="params-section-title">HOVER ARTIST LABEL</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Font Style</span>
            <select onchange="PARAMS.hoverArtistFontStyle = this.value; buildSceneParametersPanel()" style="background:#111;color:#fff;border:1px solid #333;padding:2px 4px;font-size:11px;font-family:monospace;cursor:pointer;">
                <option value="normal" ${PARAMS.hoverArtistFontStyle === 'normal' ? 'selected' : ''}>normal</option>
                <option value="bold" ${PARAMS.hoverArtistFontStyle === 'bold' ? 'selected' : ''}>bold</option>
                <option value="italic" ${PARAMS.hoverArtistFontStyle === 'italic' ? 'selected' : ''}>italic</option>
            </select>
            <span class="params-slider-value">${PARAMS.hoverArtistFontStyle}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Font Size</span>
            <input type="range" class="params-slider" min="8" max="120" step="1" value="${PARAMS.hoverArtistFontSize}" onchange="updateSceneSliderParam('hoverArtistFontSize', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverArtistFontSize')">${PARAMS.hoverArtistFontSize}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Letter Spacing</span>
            <input type="range" class="params-slider" min="-5" max="20" step="0.5" value="${PARAMS.hoverArtistLetterSpacing}" onchange="updateSceneSliderParam('hoverArtistLetterSpacing', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverArtistLetterSpacing')">${PARAMS.hoverArtistLetterSpacing}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Artist Color</span>
            <input type="color" value="${PARAMS.hoverArtistColor}" onchange="PARAMS.hoverArtistColor = this.value; buildSceneParametersPanel()" style="width:32px;height:20px;border:none;cursor:pointer;padding:0;">
            <span class="params-slider-value">${PARAMS.hoverArtistColor}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Artist Opacity</span>
            <input type="range" class="params-slider" min="0" max="1" step="0.05" value="${PARAMS.hoverArtistOpacity}" onchange="updateSceneSliderParam('hoverArtistOpacity', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'hoverArtistOpacity')">${PARAMS.hoverArtistOpacity}</span>
        </div>
        <div class="params-section-title">MOBILE HOVER LABELS</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Font Size</span>
            <input type="range" class="params-slider" min="10" max="120" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-font-size')) || 60}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-font-size', this.value+'px'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-font-size')) || 60}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Opacity</span>
            <input type="range" class="params-slider" min="0" max="1" step="0.05" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-opacity')) || 0.9}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-opacity', this.value); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-opacity')) || 0.9}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Offset Y</span>
            <input type="range" class="params-slider" min="-200" max="200" step="5" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-offset-y')) || 40}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-offset-y', this.value+'px'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-offset-y')) || 40}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Name Offset Y</span>
            <input type="range" class="params-slider" min="-200" max="200" step="5" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-name-offset-y')) || 80}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-name-offset-y', this.value+'px'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-name-offset-y')) || 80}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotation X</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-rotation-x')) || 0}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-rotation-x', this.value+'deg'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-rotation-x')) || 0}°</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotation Y</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-rotation-y')) || 90}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-rotation-y', this.value+'deg'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-rotation-y')) || 90}°</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotation Z</span>
            <input type="range" class="params-slider" min="-180" max="180" step="5" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-rotation-z')) || 0}" onchange="document.documentElement.style.setProperty('--mobile-hover-label-rotation-z', this.value+'deg'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-label-rotation-z')) || 0}°</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Artist Font Size</span>
            <input type="range" class="params-slider" min="10" max="100" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-artist-font-size')) || 35}" onchange="document.documentElement.style.setProperty('--mobile-hover-artist-font-size', this.value+'px'); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-artist-font-size')) || 35}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Artist Opacity</span>
            <input type="range" class="params-slider" min="0" max="1" step="0.05" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-artist-opacity')) || 0.7}" onchange="document.documentElement.style.setProperty('--mobile-hover-artist-opacity', this.value); buildSceneParametersPanel()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--mobile-hover-artist-opacity')) || 0.7}</span>
        </div>
        <div class="params-section-title">PANEL LAYOUT</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Right Offset</span>
            <input type="range" class="params-slider" min="0" max="100" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-right-offset')) || 20}" onchange="document.documentElement.style.setProperty('--panel-right-offset', this.value+'px'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'px'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-right-offset')) || 20}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Bottom Offset</span>
            <input type="range" class="params-slider" min="0" max="100" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-bottom-offset')) || 30}" onchange="document.documentElement.style.setProperty('--panel-bottom-offset', this.value+'px'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'px'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-bottom-offset')) || 30}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Border Radius</span>
            <input type="range" class="params-slider" min="0" max="50" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-border-radius')) || 24}" onchange="document.documentElement.style.setProperty('--panel-border-radius', this.value+'px'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'px'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-border-radius')) || 24}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">BG Opacity</span>
            <input type="range" class="params-slider" min="0" max="1" step="0.05" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-bg-opacity')) || 0.4}" onchange="document.documentElement.style.setProperty('--panel-bg-opacity', this.value); this.parentElement.querySelector('.params-slider-value').textContent=this.value">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-bg-opacity')) || 0.4}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Blur (px)</span>
            <input type="range" class="params-slider" min="0" max="30" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-blur')) || 8}" onchange="document.documentElement.style.setProperty('--panel-blur', this.value+'px'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'px'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-blur')) || 8}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Anim Duration (s)</span>
            <input type="range" class="params-slider" min="0.1" max="3" step="0.1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-animation-duration')) || 0.8}" onchange="document.documentElement.style.setProperty('--panel-animation-duration', this.value+'s'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'s'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-animation-duration')) || 0.8}s</span>
        </div>
        <div class="params-section-title">PANEL MEDIA</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Padding H (%)</span>
            <input type="range" class="params-slider" min="0" max="25" step="1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-media-padding-horizontal')) || 10}" onchange="document.documentElement.style.setProperty('--panel-media-padding-horizontal', this.value+'%'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'%'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-media-padding-horizontal')) || 10}%</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Padding Top (rem)</span>
            <input type="range" class="params-slider" min="0" max="5" step="0.25" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-media-padding-top')) || 2}" onchange="document.documentElement.style.setProperty('--panel-media-padding-top', this.value+'rem'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'rem'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-media-padding-top')) || 2}rem</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Spacing (px)</span>
            <input type="range" class="params-slider" min="0" max="50" step="2" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-media-spacing')) || 20}" onchange="document.documentElement.style.setProperty('--panel-media-spacing', this.value+'px'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'px'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-media-spacing')) || 20}px</span>
        </div>
        <div class="params-section-title">PANEL TYPOGRAPHY</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Title Font Size (rem)</span>
            <input type="range" class="params-slider" min="0.5" max="4" step="0.1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-title-font-size')) || 2}" onchange="document.documentElement.style.setProperty('--panel-title-font-size', this.value+'rem'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'rem'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-title-font-size')) || 2}rem</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Meta Font Size (rem)</span>
            <input type="range" class="params-slider" min="0.3" max="2" step="0.05" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-meta-font-size')) || 0.9}" onchange="document.documentElement.style.setProperty('--panel-meta-font-size', this.value+'rem'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'rem'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-meta-font-size')) || 0.9}rem</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Desc Font Size (rem)</span>
            <input type="range" class="params-slider" min="0.3" max="2" step="0.05" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-desc-font-size')) || 0.95}" onchange="document.documentElement.style.setProperty('--panel-desc-font-size', this.value+'rem'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'rem'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-desc-font-size')) || 0.95}rem</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Text Spacing (rem)</span>
            <input type="range" class="params-slider" min="0.5" max="4" step="0.1" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-text-spacing')) || 1.5}" onchange="document.documentElement.style.setProperty('--panel-text-spacing', this.value+'rem'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'rem'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-text-spacing')) || 1.5}rem</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Content Padding (rem)</span>
            <input type="range" class="params-slider" min="0.5" max="5" step="0.25" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-content-padding')) || 2}" onchange="document.documentElement.style.setProperty('--panel-content-padding', this.value+'rem'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'rem'">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--panel-content-padding')) || 2}rem</span>
        </div>
        <div class="params-section-title">SELECTED OBJECT ROTATION</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotate X Speed</span>
            <input type="range" class="params-slider" min="-5" max="5" step="0.1" value="${PARAMS.selectedRotateX}" onchange="updateSceneSliderParam('selectedRotateX', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'selectedRotateX')">${PARAMS.selectedRotateX}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotate Y Speed</span>
            <input type="range" class="params-slider" min="-5" max="5" step="0.1" value="${PARAMS.selectedRotateY}" onchange="updateSceneSliderParam('selectedRotateY', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'selectedRotateY')">${PARAMS.selectedRotateY}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Rotate Z Speed</span>
            <input type="range" class="params-slider" min="-5" max="5" step="0.1" value="${PARAMS.selectedRotateZ}" onchange="updateSceneSliderParam('selectedRotateZ', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'selectedRotateZ')">${PARAMS.selectedRotateZ}</span>
        </div>
        <div class="params-section-title">SELECTION & ZOOM</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Ortho Zoom Factor</span>
            <input type="range" class="params-slider" min="0.05" max="1" step="0.05" value="${PARAMS.selectionZoomFactorOrtho}" onchange="updateSceneSliderParam('selectionZoomFactorOrtho', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'selectionZoomFactorOrtho')">${PARAMS.selectionZoomFactorOrtho}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Perspective Zoom Factor</span>
            <input type="range" class="params-slider" min="0.05" max="1" step="0.05" value="${PARAMS.selectionZoomFactorPerspective}" onchange="updateSceneSliderParam('selectionZoomFactorPerspective', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'selectionZoomFactorPerspective')">${PARAMS.selectionZoomFactorPerspective}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Zoom Factor Mobile</span>
            <input type="range" class="params-slider" min="0.05" max="1" step="0.05" value="${PARAMS.selectionZoomFactorMobile}" onchange="updateSceneSliderParam('selectionZoomFactorMobile', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'selectionZoomFactorMobile')">${PARAMS.selectionZoomFactorMobile}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Zoom In Duration (ms)</span>
            <input type="range" class="params-slider" min="200" max="3000" step="100" value="${PARAMS.zoomInAnimationDuration}" onchange="updateSceneSliderParam('zoomInAnimationDuration', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'zoomInAnimationDuration')">${PARAMS.zoomInAnimationDuration}ms</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Zoom Out Duration (ms)</span>
            <input type="range" class="params-slider" min="200" max="3000" step="100" value="${PARAMS.zoomOutAnimationDuration}" onchange="updateSceneSliderParam('zoomOutAnimationDuration', this.value)">
            <span class="params-slider-value" onclick="makeSceneEditable(this, 'zoomOutAnimationDuration')">${PARAMS.zoomOutAnimationDuration}ms</span>
        </div>
        <span class="params-label">Show Price</span>
        <span class="params-value ${PARAMS.showPrice ? 'on' : 'off'}">${PARAMS.showPrice ? 'YES' : 'NO'}</span>
        <div class="params-toggle ${PARAMS.showPrice ? 'on' : ''}" onclick="PARAMS.showPrice = !PARAMS.showPrice; buildSceneParametersPanel()"></div>
        <div class="params-section-title">WIREFRAME TOGGLE</div>
        <div class="params-slider-row">
            <span class="params-slider-label">Toggle Size (px)</span>
            <input type="range" class="params-slider" min="16" max="60" step="2" value="${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--wire-toggle-size')) || 28}" onchange="document.documentElement.style.setProperty('--wire-toggle-size', this.value+'px'); this.parentElement.querySelector('.params-slider-value').textContent=this.value+'px'; updateCameraDotsVisibility(); positionMarquee()">
            <span class="params-slider-value">${parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--wire-toggle-size')) || 28}px</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Mobile Toggle Scale</span>
            <input type="range" class="params-slider" min="0.5" max="4" step="0.1" value="${PARAMS.mobileToggleScale}" onchange="PARAMS.mobileToggleScale = parseFloat(this.value); document.documentElement.style.setProperty('--mobile-toggle-scale', this.value); this.parentElement.querySelector('.params-slider-value').textContent=this.value; updateCameraDotsVisibility(); positionRightToggles(); saveParamsToStorage()">
            <span class="params-slider-value">${PARAMS.mobileToggleScale}</span>
        </div>
        <div class="params-section-title">COLORS</div>
        <div class="params-slider-row">
            <span class="params-slider-label">UI Color</span>
            <input type="color" value="${getComputedStyle(document.documentElement).getPropertyValue('--ui-color').trim() || '#000000'}" onchange="document.documentElement.style.setProperty('--ui-color', this.value); this.parentElement.querySelector('.params-slider-value').textContent=this.value" style="width:32px;height:20px;border:none;cursor:pointer;padding:0;">
            <span class="params-slider-value">${getComputedStyle(document.documentElement).getPropertyValue('--ui-color').trim() || '#000000'}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">BG Color</span>
            <input type="color" value="${getComputedStyle(document.documentElement).getPropertyValue('--bg-color').trim() || '#ffffff'}" onchange="document.documentElement.style.setProperty('--bg-color', this.value); document.getElementById('scene-container').style.background=this.value; this.parentElement.querySelector('.params-slider-value').textContent=this.value" style="width:32px;height:20px;border:none;cursor:pointer;padding:0;">
            <span class="params-slider-value">${getComputedStyle(document.documentElement).getPropertyValue('--bg-color').trim() || '#ffffff'}</span>
        </div>
        <div class="params-slider-row">
            <span class="params-slider-label">Accent Color</span>
            <input type="color" value="${getComputedStyle(document.documentElement).getPropertyValue('--accent-color').trim() || '#FF006A'}" onchange="document.documentElement.style.setProperty('--accent-color', this.value); this.parentElement.querySelector('.params-slider-value').textContent=this.value" style="width:32px;height:20px;border:none;cursor:pointer;padding:0;">
            <span class="params-slider-value">${getComputedStyle(document.documentElement).getPropertyValue('--accent-color').trim() || '#FF006A'}</span>
        </div>
    `;
    attachValueEditing(grid);
}
window.updateMarqueeParam = function(param, value, el) {
    const root = document.documentElement;
    const slider = el || (typeof event !== 'undefined' ? event.target : null);
    const valueSpan = slider ? slider.parentElement.querySelector('.params-slider-value') : null;
    if (param === 'speed') {
        root.style.setProperty('--marquee-speed', value + 's');
        if (valueSpan) valueSpan.textContent = value + 's';
    } else if (param === 'fontSize') {
        root.style.setProperty('--marquee-font-size', value + 'px');
        if (valueSpan) valueSpan.textContent = value + 'px';
    } else if (param === 'verticalGap') {
        root.style.setProperty('--marquee-vertical-gap', value + 'px');
        if (valueSpan) valueSpan.textContent = value + 'px';
    } else if (param === 'color') {
        root.style.setProperty('--marquee-color', value);
        if (valueSpan) valueSpan.textContent = value;
    } else if (param === 'mobileSpeed') {
        root.style.setProperty('--marquee-mobile-speed', value + 's');
        if (valueSpan) valueSpan.textContent = value + 's';
    } else if (param === 'mobileFontSize') {
        root.style.setProperty('--marquee-mobile-font-size', value + 'px');
        if (valueSpan) valueSpan.textContent = value + 'px';
    } else if (param === 'mobileTopOffset') {
        root.style.setProperty('--marquee-mobile-top-offset', value + 'px');
        if (valueSpan) valueSpan.textContent = value + 'px';
        positionMarquee();
    }
};
window.updateSceneSliderParam = function(paramName, value, sliderEl) {
    runtimeParams[paramName] = parseFloat(value);
    PARAMS[paramName] = parseFloat(value);
    const slider = sliderEl || (typeof event !== 'undefined' && event.target) || null;
    const valueSpan = slider ? slider.parentElement.querySelector('.params-slider-value') : null;
    if (valueSpan) {
        const pxParams = ['panelWidth', 'panelTopOffset', 'hoverLabelFontSize', 'hoverLabelOffsetY', 'hoverLabelNameOffsetY', 'hoverLiftDistance', 'hoverLabelFontBox', 'hoverArtistFontSize'];
        const degParams = ['cameraMinPolarAngle', 'cameraMaxPolarAngle', 'hoverLabelRotationX', 'hoverLabelRotationY', 'hoverLabelRotationZ'];
        const msParams = ['hoverAnimationDuration', 'zoomInAnimationDuration', 'zoomOutAnimationDuration'];
        let suffix = '';
        if (pxParams.includes(paramName)) suffix = 'px';
        else if (degParams.includes(paramName)) suffix = '°';
        else if (msParams.includes(paramName)) suffix = 'ms';
        valueSpan.textContent = value + suffix;
    }
   if (paramName === 'cameraPerspectiveFov') {
        // 1. PERSPECTIVE FOV FIX
        // Hangi kamerada olursan ol, Perspective kamerayı güncelle
        if (camera) {
            camera.fov = parseFloat(value);
            camera.updateProjectionMatrix();
        }
        markDirty();

    } else if (paramName === 'cameraOrthoFov') {
        // 2. ORTHO FOV FIX
        // Hangi kamerada olursan ol, Ortho kamerayı güncelle
        const orthoFov = parseFloat(value) || 45;
        // Eğer mobildeysek mobile zoom, değilsek desktop zoom'u baz al
        const baseZoom = currentCameraIndex === 'mobile' ? (runtimeParams.initialCameraZoomMobile || 2) : (runtimeParams.initialCameraZoom || 3);
        
        if (cameraOrtho) {
            cameraOrtho.zoom = baseZoom * (45 / orthoFov);
            cameraOrtho.updateProjectionMatrix();
        }
        markDirty();

    } else if (paramName === 'initialCameraZoom') {
        // 3. INITIAL ZOOM (DESKTOP) FIX
        const newZoom = parseFloat(value);
        // Sadece Desktop modundaysak uygula
        if (currentCameraIndex !== 'mobile') {
            const orthoFov = runtimeParams.cameraOrthoFov || 45;
            if (cameraOrtho) {
                cameraOrtho.zoom = newZoom * (45 / orthoFov);
                cameraOrtho.updateProjectionMatrix();
            }
        }
        markDirty();

    } else if (paramName === 'initialCameraZoomMobile') {
        // 4. INITIAL ZOOM (MOBILE) FIX
        const newZoomMobile = parseFloat(value);
        // Sadece Mobile modundaysak uygula
        if (currentCameraIndex === 'mobile') {
            const mobileFov = runtimeParams.camMobileFov || 45;
            if (cameraOrtho) {
                cameraOrtho.zoom = newZoomMobile * (45 / mobileFov);
                cameraOrtho.updateProjectionMatrix();
            }
        }
        markDirty();

    } else if (paramName === 'cameraMinDist') {
    } else if (paramName === 'cameraMinDist') {
        if (controls) controls.minDistance = parseFloat(value);
        if (controlsOrtho) controlsOrtho.minDistance = parseFloat(value);
    } else if (paramName === 'cameraMaxDist') {
        if (controls) controls.maxDistance = parseFloat(value);
        if (controlsOrtho) controlsOrtho.maxDistance = parseFloat(value);
    } else if (paramName.startsWith('camLook')) {
        const target = new THREE.Vector3(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        if (currentCameraMode === 'persp') {
            controls.target.copy(target);
            controls.update();
        } else {
            controlsOrtho.target.copy(target);
            controlsOrtho.update();
        }
    } else if (paramName.startsWith('cam1Pos') && currentCameraIndex === 1) {
        cameraOrtho.position.set(runtimeParams.cam1PosX, runtimeParams.cam1PosY, runtimeParams.cam1PosZ);
        cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controlsOrtho.update();
        applyLockState();
    } else if (paramName.startsWith('cam2Pos') && currentCameraIndex === 2) {
        cameraOrtho.position.set(runtimeParams.cam2PosX, runtimeParams.cam2PosY, runtimeParams.cam2PosZ);
        cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controlsOrtho.update();
        applyLockState();
    } else if (paramName.startsWith('cam3Pos') && currentCameraIndex === 3) {
        camera.position.set(runtimeParams.cam3PosX, runtimeParams.cam3PosY, runtimeParams.cam3PosZ);
        camera.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controls.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controls.update();
        applyLockState();
    } else if (paramName === 'camMobileFov') {
        // Sync CSS root variable with editor value
        document.documentElement.style.setProperty('--cam-mobile-fov', parseFloat(value));
        if (currentCameraIndex === 'mobile') {
            if (currentCameraMode === 'persp') {
                camera.fov = parseFloat(value);
                camera.updateProjectionMatrix();
            } else {
                cameraOrtho.zoom = runtimeParams.initialCameraZoomMobile * (45 / parseFloat(value));
                cameraOrtho.updateProjectionMatrix();
            }
        }
    } else if (paramName.startsWith('camMobilePos') && currentCameraIndex === 'mobile') {
        cameraOrtho.position.set(runtimeParams.camMobilePosX, runtimeParams.camMobilePosY, runtimeParams.camMobilePosZ);
        cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
        controlsOrtho.update();
    } else if (paramName === 'initialCameraZoomMobile' && currentCameraIndex === 'mobile') {
        const fovAdj = runtimeParams.camMobileFov || 45;
        cameraOrtho.zoom = runtimeParams.initialCameraZoomMobile * (45 / fovAdj);
        cameraOrtho.updateProjectionMatrix();
    } else if (paramName === 'initialCameraZoom' && currentCameraIndex !== 'mobile' && currentCameraMode === 'ortho') {
        const orthoFovAdj = runtimeParams.cameraOrthoFov || 45;
        cameraOrtho.zoom = runtimeParams.initialCameraZoom * (45 / orthoFovAdj);
        cameraOrtho.updateProjectionMatrix();
    }
    saveParamsToStorage();
};
window.makeSceneEditable = function(span, paramName) {
    const currentValue = runtimeParams[paramName];
    const parentRow = span.parentElement;
    const input = document.createElement('input');
    input.type = 'number';
    input.className = 'params-slider-input';
    input.value = currentValue;
    input.step = 'any';
    const finishEdit = () => {
        const newValue = parseFloat(input.value);
        if (!isNaN(newValue)) {
            runtimeParams[paramName] = newValue;
            PARAMS[paramName] = newValue;
            const slider = parentRow ? parentRow.querySelector('.params-slider') : null;
            if (slider) slider.value = newValue;
            if (paramName.startsWith('cam1Pos') && currentCameraIndex === 1) {
                cameraOrtho.position.set(runtimeParams.cam1PosX, runtimeParams.cam1PosY, runtimeParams.cam1PosZ);
                cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.update();
                applyLockState();
            } else if (paramName.startsWith('cam2Pos') && currentCameraIndex === 2) {
                cameraOrtho.position.set(runtimeParams.cam2PosX, runtimeParams.cam2PosY, runtimeParams.cam2PosZ);
                cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.update();
                applyLockState();
            } else if (paramName.startsWith('cam3Pos') && currentCameraIndex === 3) {
                camera.position.set(runtimeParams.cam3PosX, runtimeParams.cam3PosY, runtimeParams.cam3PosZ);
                camera.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controls.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controls.update();
                applyLockState();
            } else if (paramName.startsWith('camMobilePos') && currentCameraIndex === 'mobile') {
                cameraOrtho.position.set(runtimeParams.camMobilePosX, runtimeParams.camMobilePosY, runtimeParams.camMobilePosZ);
                cameraOrtho.lookAt(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.target.set(PARAMS.camLookX, PARAMS.camLookY, PARAMS.camLookZ);
                controlsOrtho.update();
            } else if (paramName === 'camMobileFov') {
                document.documentElement.style.setProperty('--cam-mobile-fov', newValue);
                if (currentCameraIndex === 'mobile') {
                    if (currentCameraMode === 'persp') {
                        camera.fov = newValue;
                        camera.updateProjectionMatrix();
                    } else {
                        cameraOrtho.zoom = runtimeParams.initialCameraZoomMobile * (45 / newValue);
                        cameraOrtho.updateProjectionMatrix();
                    }
                }
            } else if (paramName === 'initialCameraZoomMobile' && currentCameraIndex === 'mobile') {
                const fovAdj2 = runtimeParams.camMobileFov || 45;
                cameraOrtho.zoom = runtimeParams.initialCameraZoomMobile * (45 / fovAdj2);
                cameraOrtho.updateProjectionMatrix();
            } else if (paramName === 'initialCameraZoom' && currentCameraIndex !== 'mobile' && currentCameraMode === 'ortho') {
                const orthoFovAdj2 = runtimeParams.cameraOrthoFov || 45;
                cameraOrtho.zoom = runtimeParams.initialCameraZoom * (45 / orthoFovAdj2);
                cameraOrtho.updateProjectionMatrix();
            } else if (paramName === 'cameraOrthoFov' && currentCameraIndex !== 'mobile' && currentCameraMode === 'ortho') {
                const orthoFovEdit = newValue || 45;
                cameraOrtho.zoom = runtimeParams.initialCameraZoom * (45 / orthoFovEdit);
                cameraOrtho.updateProjectionMatrix();
            }
        }
        buildSceneParametersPanel();
    };
    input.onblur = finishEdit;
    input.onkeydown = (e) => {
        if (e.key === 'Enter') finishEdit();
        if (e.key === 'Escape') buildSceneParametersPanel();
    };
    span.replaceWith(input);
    input.focus();
    input.select();
};
window.toggleParam = function(paramName) {
    runtimeParams[paramName] = !runtimeParams[paramName];
    applyRuntimeParameters();
    buildParametersPanel();
    saveParamsToStorage();
};
window.updateSliderParam = function(paramName, value, sliderEl) {
    runtimeParams[paramName] = parseFloat(value);
    applyRuntimeParameters();
    const slider = sliderEl || (typeof event !== 'undefined' && event.target) || null;
    const valueSpan = slider ? slider.parentElement.querySelector('.params-slider-value') : null;
    if (valueSpan) {
        valueSpan.textContent = paramName === 'filterHueRotate' ? value + '°' : value;
    }
    saveParamsToStorage();
};
window.updateSliderDisplay = function(slider) {
    const valueSpan = slider.parentElement.querySelector('.params-slider-value');
    if (valueSpan) {
        valueSpan.textContent = slider.value;
    }
};
window.makeEditable = function(span, paramName) {
    const currentValue = runtimeParams[paramName];
    const parentRow = span.parentElement;
    const input = document.createElement('input');
    input.type = 'number';
    input.className = 'params-slider-input';
    input.value = currentValue;
    input.step = 'any';
    const finishEdit = () => {
        const newValue = parseFloat(input.value);
        if (!isNaN(newValue)) {
            runtimeParams[paramName] = newValue;
            applyRuntimeParameters();
            const slider = parentRow ? parentRow.querySelector('.params-slider') : null;
            if (slider) slider.value = newValue;
        }
        buildParametersPanel();
    };
    input.onblur = finishEdit;
    input.onkeydown = (e) => {
        if (e.key === 'Enter') finishEdit();
        if (e.key === 'Escape') buildParametersPanel();
    };
    span.replaceWith(input);
    input.focus();
    input.select();
};
window.switchHdriPreset = function(presetNum, forMobile) {
    if (hdriPresets && hdriPresets[presetNum]) {
        if (forMobile) {
            currentHdriPresetMobile = presetNum;
            runtimeParams.hdriPresetMobile = presetNum;
            // If currently on mobile, apply immediately
            if (isMobileView()) loadHDRI(presetNum, null, true);
        } else {
            currentHdriPreset = presetNum;
            runtimeParams.hdriPreset = presetNum;
            // If currently on desktop, apply immediately
            if (!isMobileView()) loadHDRI(presetNum, null, false);
        }
        saveParamsToStorage();
        if (editorModeActive) buildParametersPanel();
    }
};
window.updateCam3Visibility = function() {
    const btnCam3p = document.getElementById('btn-cam3');
    if (btnCam3p) {
        if (editorModeActive || runtimeParams.cam3ShowPublic) {
            btnCam3p.style.display = 'inline-block';
        } else {
            btnCam3p.style.display = 'none';
        }
    }
    updateCameraDotsVisibility();
};
window.updateCameraDotsVisibility = function() {
    const isMobileNow = isMobileView();
    const dot1 = document.getElementById('cam-dot-1');
    const dot2 = document.getElementById('cam-dot-2');
    const dot3 = document.getElementById('cam-dot-3');
    const camToggle = document.getElementById('cam-toggle');
    const wireToggle = document.getElementById('wire-toggle');
    const row1 = document.getElementById('right-toggle-row1');
    const row2 = document.getElementById('right-toggle-row2');
    const btnCam2 = document.getElementById('btn-cam2');
    const grid3dToggle = document.getElementById('grid-3d-toggle');
    if (isMobileNow) {
        // On mobile: set proper track widths so toggles align correctly
        const mSize = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--wire-toggle-size')) || 12;
        const mScale = PARAMS.mobileToggleScale || 1;
        const scaled = mSize * mScale;
        const mTrackW = 2 * scaled + 0.25 * scaled;
        const mWireTrack = wireToggle ? wireToggle.querySelector('.wire-toggle-track') : null;
        if (mWireTrack) {
            mWireTrack.style.width = mTrackW + 'px';
            mWireTrack.style.height = scaled + 'px';
            mWireTrack.style.borderRadius = (scaled / 2) + 'px';
            const mWireThumbActive = mTrackW - scaled;
            mWireTrack.dataset.thumbActiveLeft = mWireThumbActive + 'px';
            const mWireThumb = mWireTrack.querySelector('.wire-toggle-thumb');
            if (mWireThumb) {
                mWireThumb.style.width = scaled + 'px';
                mWireThumb.style.height = scaled + 'px';
                if (wireToggle.classList.contains('active')) {
                    mWireThumb.style.left = mWireThumbActive + 'px';
                }
            }
        }
        if (grid3dToggle) {
            const mG3dTrack = grid3dToggle.querySelector('.grid-3d-track');
            const mG3dIcons = grid3dToggle.querySelector('.grid-3d-icons');
            if (mG3dTrack) {
                mG3dTrack.style.width = mTrackW + 'px';
                mG3dTrack.style.height = scaled + 'px';
                mG3dTrack.style.borderRadius = (scaled / 2) + 'px';
                const mG3dThumbActive = mTrackW - scaled;
                mG3dTrack.dataset.thumbActiveLeft = mG3dThumbActive + 'px';
                const mG3dThumb = mG3dTrack.querySelector('.grid-3d-thumb');
                if (mG3dThumb) {
                    mG3dThumb.style.width = scaled + 'px';
                    mG3dThumb.style.height = scaled + 'px';
                    if (grid3dToggle.classList.contains('active')) {
                        mG3dThumb.style.left = mG3dThumbActive + 'px';
                    }
                }
            }
            if (mG3dIcons) {
                // Icons stay at base size, padded to align centers with thumb positions
                const iconSize = mSize * 0.65;
                const padLR = Math.max(0, (scaled - iconSize) / 2);
                mG3dIcons.style.width = mTrackW + 'px';
                mG3dIcons.style.boxSizing = 'border-box';
                mG3dIcons.style.paddingLeft = padLR + 'px';
                mG3dIcons.style.paddingRight = padLR + 'px';
            }
        }
        positionRightToggles();
        return;
    }
    if (dot1) dot1.style.display = '';
    if (dot2) dot2.style.display = runtimeParams.cam2ShowPublic ? '' : 'none';
    if (dot3) dot3.style.display = runtimeParams.cam3ShowPublic ? '' : 'none';
    if (btnCam2) btnCam2.style.display = runtimeParams.cam2ShowPublic || editorModeActive ? '' : 'none';
    let visibleCount = 1;
    if (runtimeParams.cam2ShowPublic) visibleCount++;
    if (runtimeParams.cam3ShowPublic) visibleCount++;
    // row2 = camera dots row now
    if (visibleCount <= 1) {
        if (camToggle) camToggle.style.display = 'none';
        if (!gridModeActive && row2) row2.style.display = 'none';
    } else {
        if (camToggle) camToggle.style.display = '';
        if (!gridModeActive && row2) row2.style.display = 'flex';
    }
    // Calculate unified width: max of grid-3d track (2 positions) and cam-dots
    const size = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--wire-toggle-size')) || 12;
    const grid3dNatural = 2 * size + 0.25 * size; // track = 2 stops + 1 gap
    let camWidth;
    if (visibleCount > 1) {
        camWidth = visibleCount * size + (visibleCount - 1) * size * 0.25;
    } else {
        camWidth = size * 1.8;
    }
    const unifiedWidth = Math.max(grid3dNatural, camWidth);
    // Set wire toggle track width (row1)
    const wireTrack = wireToggle ? wireToggle.querySelector('.wire-toggle-track') : null;
    if (wireTrack) {
        wireTrack.style.width = unifiedWidth + 'px';
        const wireThumbActive = visibleCount > 1
            ? (visibleCount - 1) * 1.25 * size
            : (unifiedWidth - size);
        wireTrack.dataset.thumbActiveLeft = wireThumbActive + 'px';
        const wireThumb = wireTrack.querySelector('.wire-toggle-thumb');
        if (wireThumb && wireToggle.classList.contains('active')) {
            wireThumb.style.left = wireThumbActive + 'px';
        }
    }
    // Set grid-3d-toggle track width (row0)
    if (grid3dToggle) {
        const g3dTrack = grid3dToggle.querySelector('.grid-3d-track');
        const g3dIcons = grid3dToggle.querySelector('.grid-3d-icons');
        if (g3dTrack) {
            g3dTrack.style.width = unifiedWidth + 'px';
            const g3dThumbActive = unifiedWidth - size;
            g3dTrack.dataset.thumbActiveLeft = g3dThumbActive + 'px';
            const g3dThumb = g3dTrack.querySelector('.grid-3d-thumb');
            if (g3dThumb && grid3dToggle.classList.contains('active')) {
                g3dThumb.style.left = g3dThumbActive + 'px';
            }
        }
        if (g3dIcons) {
            g3dIcons.style.width = unifiedWidth + 'px';
        }
    }
    positionRightToggles();
};
function applyRuntimeParameters() {
    markDirty();
    const isMobile = isMobileView();
    // Apply FOV: mobile uses camMobileFov, desktop ortho uses cameraOrthoFov
    if (currentCameraIndex === 'mobile') {
        const mobileFov = runtimeParams.camMobileFov || 45;
        if (currentCameraMode === 'persp' && camera) {
            camera.fov = mobileFov;
            camera.updateProjectionMatrix();
        } else if (currentCameraMode === 'ortho' && cameraOrtho) {
            cameraOrtho.zoom = runtimeParams.initialCameraZoomMobile * (45 / mobileFov);
            cameraOrtho.updateProjectionMatrix();
        }
    } else if (currentCameraMode === 'ortho' && cameraOrtho) {
        const orthoFov = runtimeParams.cameraOrthoFov || 45;
        cameraOrtho.zoom = runtimeParams.initialCameraZoom * (45 / orthoFov);
        cameraOrtho.updateProjectionMatrix();
    }
    if (lightHemi) lightHemi.intensity = runtimeParams.lightHemisphereIntensity;
    if (lightSun) {
        lightSun.intensity = runtimeParams.lightSunIntensity;
        lightSun.position.set(runtimeParams.lightSunPosX, runtimeParams.lightSunPosY, runtimeParams.lightSunPosZ);
        const useShadows = isMobile ? runtimeParams.shadowsEnabledMobile : runtimeParams.shadowsEnabledDesktop;
        lightSun.castShadow = useShadows;
    }
    if (lightFill) { lightFill.intensity = runtimeParams.lightFillIntensity; lightFill.position.set(runtimeParams.lightFillPosX, runtimeParams.lightFillPosY, runtimeParams.lightFillPosZ); }
    if (lightRim) { lightRim.intensity = runtimeParams.lightRimIntensity; lightRim.position.set(runtimeParams.lightRimPosX, runtimeParams.lightRimPosY, runtimeParams.lightRimPosZ); }
    if (lightAccent) { lightAccent.intensity = runtimeParams.lightAccentIntensity; lightAccent.position.set(runtimeParams.lightAccentPosX, runtimeParams.lightAccentPosY, runtimeParams.lightAccentPosZ); }
    const useRectAreaLight = isMobile ? runtimeParams.rectarealightEnabledMobile : runtimeParams.rectarealightEnabledDesktop;
    if (rectLight1) {
        rectLight1.visible = useRectAreaLight;
        rectLight1.intensity = runtimeParams.rectarealightIntensity;
    }
    if (rectLight2) {
        rectLight2.visible = useRectAreaLight;
        rectLight2.intensity = runtimeParams.rectarealightIntensity * 0.6;
    }
    if (ssaoPass) {
        const useSSAO = isMobile ? runtimeParams.ssaoEnabledMobile : runtimeParams.ssaoEnabledDesktop;
        ssaoPass.enabled = useSSAO;
        ssaoPass.kernelRadius = runtimeParams.ssaoKernelRadius;
        ssaoPass.minDistance = 0.001;
        ssaoPass.maxDistance = 0.05 + runtimeParams.ssaoIntensity * 0.2;
    }
    if (bloomPass) {
        const useBloom = isMobile ? runtimeParams.bloomEnabledMobile : runtimeParams.bloomEnabledDesktop;
        bloomPass.enabled = useBloom;
        bloomPass.strength = runtimeParams.bloomStrength;
        bloomPass.radius = runtimeParams.bloomRadius;
        bloomPass.threshold = runtimeParams.bloomThreshold;
    }
    if (renderer) {
        const useTonemapping = isMobile ? runtimeParams.tonemappingEnabledMobile : runtimeParams.tonemappingEnabledDesktop;
        if (useTonemapping) {
            renderer.toneMapping = THREE.ACESFilmicToneMapping;
            renderer.toneMappingExposure = isMobile ? runtimeParams.tonemappingExposureMobile : runtimeParams.tonemappingExposureDesktop;
        } else {
            renderer.toneMapping = THREE.NoToneMapping;
            renderer.toneMappingExposure = 1.0;
        }
        const useShadows = isMobile ? runtimeParams.shadowsEnabledMobile : runtimeParams.shadowsEnabledDesktop;
        renderer.shadowMap.enabled = useShadows;
    }
    if (fxaaPass) {
        const useFXAA = isMobile ? runtimeParams.fxaaEnabledMobile : runtimeParams.fxaaEnabledDesktop;
        fxaaPass.enabled = useFXAA;
    }
    const useEnvMap = isMobile ? runtimeParams.envmapEnabledMobile : runtimeParams.envmapEnabledDesktop;
    if (scene) {
        scene.environment = useEnvMap ? globalEnvMap : null;
    }
    const envIntensity = isMobile ? runtimeParams.envmapIntensityMobile : runtimeParams.envmapIntensityDesktop;
    const roughnessBoost = runtimeParams.pbrRoughnessBoost || 0;
    const metalnessBoost = runtimeParams.pbrMetalnessBoost || 0;
    const useAnisotropy = isMobile ? runtimeParams.anisotropyEnabledMobile : runtimeParams.anisotropyEnabledDesktop;
    const maxAniso = renderer ? renderer.capabilities.getMaxAnisotropy() : 1;
    clickableObjects.forEach(obj => {
        obj.traverse(child => {
            if (child.isMesh && child.material && child.material.isMeshStandardMaterial) {
                const mat = child.material;
                mat.envMapIntensity = envIntensity;
                if (mat._originalRoughness === undefined) {
                    mat._originalRoughness = mat.roughness;
                }
                if (mat._originalMetalness === undefined) {
                    mat._originalMetalness = mat.metalness;
                }
                mat.roughness = Math.max(0, Math.min(1, mat._originalRoughness + roughnessBoost));
                mat.metalness = Math.max(0, Math.min(1, mat._originalMetalness + metalnessBoost));
                if (mat.map) mat.map.anisotropy = useAnisotropy ? maxAniso : 1;
                if (mat.normalMap) mat.normalMap.anisotropy = useAnisotropy ? maxAniso : 1;
                if (mat.roughnessMap) mat.roughnessMap.anisotropy = useAnisotropy ? maxAniso : 1;
                mat.needsUpdate = true;
            }
        });
    });
    if (controls) {
        controls.minPolarAngle = (90 - runtimeParams.cameraMaxPolarAngle) * Math.PI / 180;
        controls.maxPolarAngle = (90 - runtimeParams.cameraMinPolarAngle) * Math.PI / 180;
    }
    if (controlsOrtho) {
        controlsOrtho.minPolarAngle = (90 - runtimeParams.cameraMaxPolarAngle) * Math.PI / 180;
        controlsOrtho.maxPolarAngle = (90 - runtimeParams.cameraMinPolarAngle) * Math.PI / 180;
    }
    applyColorFilters();
}
function applyColorFilters() {
    if (!renderer) return;
    const canvas = renderer.domElement;
    const brightness = runtimeParams.filterBrightness;
    const contrast = runtimeParams.filterContrast;
    const saturation = runtimeParams.filterSaturation;
    const hueRotate = runtimeParams.filterHueRotate;
    canvas.style.filter = `brightness(${brightness}) contrast(${contrast}) saturate(${saturation}) hue-rotate(${hueRotate}deg)`;
}
window.resetParameters = function() {
    runtimeParams = { ...RENDER_PARAMS };
    window.runtimeParams = runtimeParams;
    applyRuntimeParameters();
    buildParametersPanel();
    clearSavedParams();
    markDirty();
};
window.resetSceneParameters = function() {
    const SCENE_PARAMS_KEYS = ['cam1PosX','cam1PosY','cam1PosZ','cam1LockMode','cam2PosX','cam2PosY','cam2PosZ','cam2LockMode','cam3PosX','cam3PosY','cam3PosZ','cam3LockMode','camMobilePosX','camMobilePosY','camMobilePosZ','camMobileLockMode','camMobileFov','initialCameraZoom','initialCameraZoomMobile','cameraLockShowPublic','cameraDefaultLocked','zoomRotateEnabled','cam2ShowPublic','cam3ShowPublic','mobileLabelsShowPublic','cameraMinPolarAngle','cameraMaxPolarAngle','envOption','panelWidth','panelTopOffset'];
    SCENE_PARAMS_KEYS.forEach(k => {
        if (RENDER_PARAMS[k] !== undefined) runtimeParams[k] = RENDER_PARAMS[k];
    });
    const SCENE_PARAMS_DEFAULTS = {};
    for (const k in PARAMS) { SCENE_PARAMS_DEFAULTS[k] = PARAMS[k]; }
    buildSceneParametersPanel();
    clearSavedParams();
    markDirty();
};
window.printParameters = function() {
    let output = "/* PANEL 2 - RENDER */\n\n";
    output += "/* SSAO */\n";
    output += `--ssao-enabled-desktop: ${runtimeParams.ssaoEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--ssao-enabled-mobile: ${runtimeParams.ssaoEnabledMobile ? 'yes' : 'no'};\n`;
    output += `--ssao-kernel-radius: ${runtimeParams.ssaoKernelRadius};\n`;
    output += `--ssao-intensity: ${runtimeParams.ssaoIntensity};\n\n`;
    output += "/* BLOOM */\n";
    output += `--bloom-enabled-desktop: ${runtimeParams.bloomEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--bloom-enabled-mobile: ${runtimeParams.bloomEnabledMobile ? 'yes' : 'no'};\n`;
    output += `--bloom-strength: ${runtimeParams.bloomStrength};\n`;
    output += `--bloom-threshold: ${runtimeParams.bloomThreshold};\n`;
    output += `--bloom-radius: ${runtimeParams.bloomRadius};\n\n`;
    output += "/* TONE MAPPING */\n";
    output += `--tonemapping-enabled-desktop: ${runtimeParams.tonemappingEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--tonemapping-enabled-mobile: ${runtimeParams.tonemappingEnabledMobile ? 'yes' : 'no'};\n`;
    output += `--tonemapping-exposure-desktop: ${runtimeParams.tonemappingExposureDesktop};\n`;
    output += `--tonemapping-exposure-mobile: ${runtimeParams.tonemappingExposureMobile};\n\n`;
    output += "/* SHADOWS */\n";
    output += `--shadows-enabled-desktop: ${runtimeParams.shadowsEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--shadows-enabled-mobile: ${runtimeParams.shadowsEnabledMobile ? 'yes' : 'no'};\n`;
    output += `--shadow-type: ${runtimeParams.shadowType};\n`;
    output += `--shadow-bias: ${runtimeParams.shadowBias};\n`;
    output += `--shadow-normal-bias: ${runtimeParams.shadowNormalBias};\n`;
    output += `--shadow-map-size-desktop: ${runtimeParams.shadowMapSizeDesktop};\n`;
    output += `--shadow-map-size-mobile: ${runtimeParams.shadowMapSizeMobile};\n\n`;
    output += "/* ANTI-ALIASING */\n";
    output += `--fxaa-enabled-desktop: ${runtimeParams.fxaaEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--fxaa-enabled-mobile: ${runtimeParams.fxaaEnabledMobile ? 'yes' : 'no'};\n\n`;
    output += "/* ANISOTROPY */\n";
    output += `--anisotropy-enabled-desktop: ${runtimeParams.anisotropyEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--anisotropy-enabled-mobile: ${runtimeParams.anisotropyEnabledMobile ? 'yes' : 'no'};\n\n`;
    output += "/* ENVIRONMENT MAP */\n";
    output += `--envmap-enabled-desktop: ${runtimeParams.envmapEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--envmap-enabled-mobile: ${runtimeParams.envmapEnabledMobile ? 'yes' : 'no'};\n`;
    output += `--envmap-intensity-desktop: ${runtimeParams.envmapIntensityDesktop};\n`;
    output += `--envmap-intensity-mobile: ${runtimeParams.envmapIntensityMobile};\n\n`;
    output += "/* RECT AREA LIGHT */\n";
    output += `--rectarealight-enabled-desktop: ${runtimeParams.rectarealightEnabledDesktop ? 'yes' : 'no'};\n`;
    output += `--rectarealight-enabled-mobile: ${runtimeParams.rectarealightEnabledMobile ? 'yes' : 'no'};\n`;
    output += `--rectarealight-intensity: ${runtimeParams.rectarealightIntensity};\n\n`;
    output += "/* LIGHTING */\n";
    output += `--light-hemisphere-intensity: ${runtimeParams.lightHemisphereIntensity};\n`;
    output += `--light-sun-intensity: ${runtimeParams.lightSunIntensity};\n`;
    output += `--light-sun-pos-x: ${runtimeParams.lightSunPosX};\n`;
    output += `--light-sun-pos-y: ${runtimeParams.lightSunPosY};\n`;
    output += `--light-sun-pos-z: ${runtimeParams.lightSunPosZ};\n`;
    output += `--light-fill-intensity: ${runtimeParams.lightFillIntensity};\n`;
    output += `--light-fill-pos-x: ${runtimeParams.lightFillPosX};\n`;
    output += `--light-fill-pos-y: ${runtimeParams.lightFillPosY};\n`;
    output += `--light-fill-pos-z: ${runtimeParams.lightFillPosZ};\n`;
    output += `--light-rim-intensity: ${runtimeParams.lightRimIntensity};\n`;
    output += `--light-rim-pos-x: ${runtimeParams.lightRimPosX};\n`;
    output += `--light-rim-pos-y: ${runtimeParams.lightRimPosY};\n`;
    output += `--light-rim-pos-z: ${runtimeParams.lightRimPosZ};\n`;
    output += `--light-accent-intensity: ${runtimeParams.lightAccentIntensity};\n`;
    output += `--light-accent-pos-x: ${runtimeParams.lightAccentPosX};\n`;
    output += `--light-accent-pos-y: ${runtimeParams.lightAccentPosY};\n`;
    output += `--light-accent-pos-z: ${runtimeParams.lightAccentPosZ};\n\n`;
    output += "/* PBR MATERIALS */\n";
    output += `--pbr-envmap-intensity: ${runtimeParams.pbrEnvmapIntensity};\n`;
    output += `--pbr-roughness-boost: ${runtimeParams.pbrRoughnessBoost};\n`;
    output += `--pbr-metalness-boost: ${runtimeParams.pbrMetalnessBoost};\n\n`;
    output += "/* COLOR CORRECTION */\n";
    output += `--filter-brightness: ${runtimeParams.filterBrightness};\n`;
    output += `--filter-contrast: ${runtimeParams.filterContrast};\n`;
    output += `--filter-saturation: ${runtimeParams.filterSaturation};\n`;
    output += `--filter-hue-rotate: ${runtimeParams.filterHueRotate};\n\n`;
    output += "/* HDRI PRESET */\n";
    output += `--hdri-preset: ${currentHdriPreset};\n`;
    output += `--hdri-preset-mobile: ${currentHdriPresetMobile};\n`;
    navigator.clipboard.writeText(output).then(() => {
        alert('PANEL 2 (RENDER) COPIED!\n\nPaste into your :root CSS section.');
    }).catch(err => {
        console.log(output);
        alert('Copy failed. Check console for data.');
    });
};
window.printSceneParameters = function() {
    const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
    let output = "/* PANEL 3 - SCENE */\n\n";
    output += "/* ENVIRONMENT OPTION */\n";
    output += `--env-option: ${runtimeParams.envOption || 1};\n\n`;
    output += "/* CAMERA 1 POSITION */\n";
    output += `--cam1-pos-x: ${runtimeParams.cam1PosX};\n`;
    output += `--cam1-pos-y: ${runtimeParams.cam1PosY};\n`;
    output += `--cam1-pos-z: ${runtimeParams.cam1PosZ};\n`;
    output += `--cam1-lock-mode: ${runtimeParams.cam1LockMode || 'none'};\n`;
    output += `--cam2-show-public: ${runtimeParams.cam2ShowPublic ? 'yes' : 'no'};\n\n`;
    output += "/* CAMERA 2 POSITION */\n";
    output += `--cam2-pos-x: ${runtimeParams.cam2PosX};\n`;
    output += `--cam2-pos-y: ${runtimeParams.cam2PosY};\n`;
    output += `--cam2-pos-z: ${runtimeParams.cam2PosZ};\n`;
    output += `--cam2-lock-mode: ${runtimeParams.cam2LockMode || 'none'};\n\n`;
    output += "/* CAMERA 3 POSITION */\n";
    output += `--cam3-pos-x: ${runtimeParams.cam3PosX};\n`;
    output += `--cam3-pos-y: ${runtimeParams.cam3PosY};\n`;
    output += `--cam3-pos-z: ${runtimeParams.cam3PosZ};\n`;
    output += `--cam3-lock-mode: ${runtimeParams.cam3LockMode || 'none'};\n`;
    output += `--cam3-show-public: ${runtimeParams.cam3ShowPublic ? 'yes' : 'no'};\n`;
    output += `--mobile-labels-show-public: ${runtimeParams.mobileLabelsShowPublic ? 'yes' : 'no'};\n\n`;
    output += "/* MOBILE CAMERA */\n";
    output += `--cam-mobile-pos-x: ${runtimeParams.camMobilePosX};\n`;
    output += `--cam-mobile-pos-y: ${runtimeParams.camMobilePosY};\n`;
    output += `--cam-mobile-pos-z: ${runtimeParams.camMobilePosZ};\n`;
    output += `--cam-mobile-lock-mode: ${runtimeParams.camMobileLockMode || 'none'};\n`;
    output += `--cam-mobile-fov: ${runtimeParams.camMobileFov};\n\n`;
    output += "/* CAMERA SETTINGS */\n";
    output += `--initial-camera-zoom: ${runtimeParams.initialCameraZoom};\n`;
    output += `--initial-camera-zoom-mobile: ${runtimeParams.initialCameraZoomMobile};\n`;
    output += `--camera-lock-show-public: ${runtimeParams.cameraLockShowPublic ? 'yes' : 'no'};\n`;
    output += `--camera-default-locked: ${runtimeParams.cameraDefaultLocked ? 'yes' : 'no'};\n`;
    output += `--zoom-rotate-enabled: ${runtimeParams.zoomRotateEnabled ? 'yes' : 'no'};\n\n`;
    output += "/* HOVER LABEL (Desktop) */\n";
    output += `--hover-label-font-size: ${runtimeParams.hoverLabelFontSize}px;\n`;
    output += `--hover-label-offset-y: ${runtimeParams.hoverLabelOffsetY}px;\n`;
    output += `--hover-label-name-offset-y: ${runtimeParams.hoverLabelNameOffsetY}px;\n\n`;
    output += "/* PANEL SETTINGS */\n";
    output += `--panel-width: ${runtimeParams.panelWidth}px;\n`;
    output += `--panel-top-offset: ${runtimeParams.panelTopOffset}px;\n\n`;
    output += "/* MARQUEE TEXT */\n";
    output += `--marquee-speed: ${css('--marquee-speed') || '250s'};\n`;
    output += `--marquee-font-size: ${css('--marquee-font-size') || '16px'};\n`;
    output += `--marquee-vertical-gap: ${css('--marquee-vertical-gap') || '-2px'};\n`;
    output += `--marquee-color: ${css('--marquee-color') || '#b6a6d0'};\n\n`;
    output += "/* MARQUEE MOBILE */\n";
    output += `--marquee-mobile-enabled: ${mobileMarqueeEnabled ? 'yes' : 'no'};\n`;
    output += `--marquee-mobile-speed: ${css('--marquee-mobile-speed') || '250s'};\n`;
    output += `--marquee-mobile-font-size: ${css('--marquee-mobile-font-size') || '7px'};\n`;
    output += `--marquee-mobile-top-offset: ${css('--marquee-mobile-top-offset') || '16px'};\n\n`;
    output += "/* CAMERA LIMITS */\n";
    output += `--camera-min-polar-angle: ${runtimeParams.cameraMinPolarAngle};\n`;
    output += `--camera-max-polar-angle: ${runtimeParams.cameraMaxPolarAngle};\n\n`;
    output += "/* CAMERA GENERAL */\n";
    output += `--camera-perspective-fov: ${PARAMS.cameraPerspectiveFov};\n`;
    output += `--camera-ortho-fov: ${PARAMS.cameraOrthoFov};\n`;
    output += `--camera-min-distance: ${PARAMS.cameraMinDist};\n`;
    output += `--camera-max-distance: ${PARAMS.cameraMaxDist};\n`;
    output += `--enable-pan: ${PARAMS.enablePan ? 'yes' : 'no'};\n`;
    output += `--mobile-pan-enabled: ${PARAMS.mobilePanEnabled ? 'yes' : 'no'};\n\n`;
    output += "/* CAM LOOK-AT */\n";
    output += `--cam-lookat-x: ${PARAMS.camLookX};\n`;
    output += `--cam-lookat-y: ${PARAMS.camLookY};\n`;
    output += `--cam-lookat-z: ${PARAMS.camLookZ};\n\n`;
    output += "/* HOVER ANIMATION */\n";
    output += `--hover-while-selected-factor: ${PARAMS.hoverWhileSelectedFactor};\n`;
    output += `--hover-lift-distance: ${PARAMS.hoverLiftDistance}px;\n`;
    output += `--hover-animation-duration: ${PARAMS.hoverAnimationDuration / 1000}s;\n\n`;
    output += "/* HOVER LABEL STYLE */\n";
    output += `--hover-label-font-style: ${PARAMS.hoverLabelFontStyle};\n`;
    output += `--hover-label-letter-spacing: ${PARAMS.hoverLabelLetterSpacing};\n`;
    output += `--hover-label-color: ${PARAMS.hoverLabelColor};\n`;
    output += `--hover-label-opacity: ${PARAMS.hoverLabelOpacity};\n`;
    output += `--hover-label-font-box: ${PARAMS.hoverLabelFontBox}px;\n`;
    output += `--hover-label-rotation-x: ${PARAMS.hoverLabelRotationX}deg;\n`;
    output += `--hover-label-rotation-y: ${PARAMS.hoverLabelRotationY}deg;\n`;
    output += `--hover-label-rotation-z: ${PARAMS.hoverLabelRotationZ}deg;\n\n`;
    output += "/* HOVER ARTIST LABEL */\n";
    output += `--hover-artist-font-size: ${PARAMS.hoverArtistFontSize}px;\n`;
    output += `--hover-artist-font-style: ${PARAMS.hoverArtistFontStyle};\n`;
    output += `--hover-artist-letter-spacing: ${PARAMS.hoverArtistLetterSpacing};\n`;
    output += `--hover-artist-color: ${PARAMS.hoverArtistColor};\n`;
    output += `--hover-artist-opacity: ${PARAMS.hoverArtistOpacity};\n\n`;
    output += "/* MOBILE HOVER LABELS */\n";
    output += `--mobile-hover-label-font-size: ${css('--mobile-hover-label-font-size') || '60px'};\n`;
    output += `--mobile-hover-label-opacity: ${css('--mobile-hover-label-opacity') || '0.9'};\n`;
    output += `--mobile-hover-label-offset-y: ${css('--mobile-hover-label-offset-y') || '40px'};\n`;
    output += `--mobile-hover-label-name-offset-y: ${css('--mobile-hover-label-name-offset-y') || '80px'};\n`;
    output += `--mobile-hover-label-rotation-x: ${css('--mobile-hover-label-rotation-x') || '0deg'};\n`;
    output += `--mobile-hover-label-rotation-y: ${css('--mobile-hover-label-rotation-y') || '90deg'};\n`;
    output += `--mobile-hover-label-rotation-z: ${css('--mobile-hover-label-rotation-z') || '0deg'};\n`;
    output += `--mobile-hover-artist-font-size: ${css('--mobile-hover-artist-font-size') || '35px'};\n`;
    output += `--mobile-hover-artist-opacity: ${css('--mobile-hover-artist-opacity') || '0.7'};\n\n`;
    output += "/* PANEL LAYOUT */\n";
    output += `--panel-right-offset: ${css('--panel-right-offset') || '20px'};\n`;
    output += `--panel-bottom-offset: ${css('--panel-bottom-offset') || '30px'};\n`;
    output += `--panel-border-radius: ${css('--panel-border-radius') || '24px'};\n`;
    output += `--panel-bg-opacity: ${css('--panel-bg-opacity') || '0.4'};\n`;
    output += `--panel-blur: ${css('--panel-blur') || '8px'};\n`;
    output += `--panel-animation-duration: ${css('--panel-animation-duration') || '0.8s'};\n\n`;
    output += "/* PANEL MEDIA */\n";
    output += `--panel-media-padding-horizontal: ${css('--panel-media-padding-horizontal') || '10%'};\n`;
    output += `--panel-media-padding-top: ${css('--panel-media-padding-top') || '2rem'};\n`;
    output += `--panel-media-spacing: ${css('--panel-media-spacing') || '20px'};\n\n`;
    output += "/* PANEL TYPOGRAPHY */\n";
    output += `--panel-title-font-size: ${css('--panel-title-font-size') || '2rem'};\n`;
    output += `--panel-meta-font-size: ${css('--panel-meta-font-size') || '0.9rem'};\n`;
    output += `--panel-desc-font-size: ${css('--panel-desc-font-size') || '0.95rem'};\n`;
    output += `--panel-text-spacing: ${css('--panel-text-spacing') || '1.5rem'};\n`;
    output += `--panel-content-padding: ${css('--panel-content-padding') || '2rem'};\n\n`;
    output += "/* SELECTED OBJECT ROTATION */\n";
    output += `--selected-rotate-x: ${PARAMS.selectedRotateX};\n`;
    output += `--selected-rotate-y: ${PARAMS.selectedRotateY};\n`;
    output += `--selected-rotate-z: ${PARAMS.selectedRotateZ};\n\n`;
    output += "/* SELECTION & ZOOM */\n";
    output += `--selection-zoom-factor-ortho: ${PARAMS.selectionZoomFactorOrtho};\n`;
    output += `--selection-zoom-factor-perspective: ${PARAMS.selectionZoomFactorPerspective};\n`;
    output += `--selection-zoom-factor-mobile: ${PARAMS.selectionZoomFactorMobile};\n`;
    output += `--zoom-in-animation-duration: ${PARAMS.zoomInAnimationDuration / 1000}s;\n`;
    output += `--zoom-out-animation-duration: ${PARAMS.zoomOutAnimationDuration / 1000}s;\n`;
    output += `--show-price: ${PARAMS.showPrice ? 'yes' : 'no'};\n\n`;
    output += "/* WIREFRAME TOGGLE */\n";
    output += `--wire-toggle-size: ${css('--wire-toggle-size') || '12px'};\n`;
    output += `--mobile-toggle-scale: ${PARAMS.mobileToggleScale};\n\n`;
    output += "/* COLORS */\n";
    output += `--ui-color: ${css('--ui-color') || '#000000'};\n`;
    output += `--bg-color: ${css('--bg-color') || '#ffffff'};\n`;
    output += `--accent-color: ${css('--accent-color') || '#FF006A'};\n`;
    navigator.clipboard.writeText(output).then(() => {
        alert('PANEL 3 (SCENE) COPIED!\n\nPaste into your :root CSS section.');
    }).catch(err => {
        console.log(output);
        alert('Copy failed. Check console for data.');
    });
};
window.selectEnvOption = function(optNum) {
    runtimeParams.envOption = optNum;
    const envObj = sceneObjects.find(o => o.modalContent && o.modalContent.name === 'Environment');
    if (!envObj || !envObj.envOptions) return;
    const opt = envObj.envOptions[optNum - 1];
    if (!opt || !opt.src) return;
    const envIndex = sceneObjects.indexOf(envObj);
    const oldMesh = loadedObjectsMap.get(envIndex);
    if (oldMesh) {
        if (selectedEditorObject === oldMesh) {
            transformControls.detach();
            selectedEditorObject = null;
        }
        scene.remove(oldMesh);
        const oldIdx = clickableObjects.indexOf(oldMesh);
        if (oldIdx !== -1) clickableObjects.splice(oldIdx, 1);
        oldMesh.traverse(c => { if (c.geometry) c.geometry.dispose(); if (c.material) { if (Array.isArray(c.material)) c.material.forEach(m => m.dispose()); else c.material.dispose(); } });
    }
    envObj.src = opt.src;
    const loader = new GLTFLoader();
    const draco = new DRACOLoader();
    draco.setDecoderPath(DRACO_DECODER_PATH);
    loader.setDRACOLoader(draco);
    loader.setMeshoptDecoder(MeshoptDecoder);
    loader.load(opt.src, (gltf) => {
        const m = gltf.scene;
        m.position.set(envObj.position.x, envObj.position.y || 0, envObj.position.z);
        if (envObj.rotation) m.rotation.set(envObj.rotation.x * Math.PI / 180, envObj.rotation.y * Math.PI / 180, envObj.rotation.z * Math.PI / 180);
        m.scale.setScalar(envObj.scale);
        m.userData = { modalContent: envObj.modalContent, clickable: false, includeInFit: true, originalPosition: m.position.clone(), sceneObjectIndex: envIndex, originalScale: envObj.scale };
        m.traverse(c => { if (c.isMesh) { c.castShadow = true; c.receiveShadow = true; } });
        scene.add(m);
        clickableObjects.push(m);
        loadedObjectsMap.set(envIndex, m);
        buildEditorPanel();
    }, undefined, (err) => { console.error("Error loading env option:", opt.src, err); });
    buildSceneParametersPanel();
};
function saveUndoState() {
    const state = {};
    loadedObjectsMap.forEach((obj, index) => {
        state[index] = {
            position: obj.position.clone(),
            rotation: obj.rotation.clone(),
            scale: obj.scale.x
        };
    });
    undoStack.push(state);
    if (undoStack.length > MAX_UNDO_STEPS) undoStack.shift();
    redoStack = [];
}
function handleKeyDown(e) {
    const tag = document.activeElement?.tagName;
    const isEditing = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
    if (isEditing) {
        if (e.key === 'Escape') { document.activeElement.blur(); return; }
        if (e.key === 'Tab') {
            const inputs = Array.from(document.querySelectorAll('#editor-panel input:not([disabled])'));
            const idx = inputs.indexOf(document.activeElement);
            if (idx >= 0) {
                e.preventDefault();
                const next = e.shiftKey ? inputs[idx - 1] : inputs[idx + 1];
                if (next) { next.focus(); next.select(); }
            }
            return;
        }
        if (e.key === 'Enter') { document.activeElement.blur(); return; }
        return;
    }
    if (e.key === 'Escape') { closePanel(); return; }
    if (e.shiftKey && (e.key === 'E' || e.key === 'e')) {
        e.preventDefault();
        toggleEditorMode();
        return;
    }
    if (e.shiftKey && (e.key === 'B' || e.key === 'b')) {
        e.preventDefault();
        toggleBoundingBoxes();
        return;
    }
    if (!editorModeActive) return;
    if (e.ctrlKey && e.key === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
    if ((e.ctrlKey && e.key === 'y') || (e.ctrlKey && e.shiftKey && e.key === 'z')) { e.preventDefault(); redo(); }
}
function undo() {
    if (undoStack.length === 0) return;
    const currentState = {};
    loadedObjectsMap.forEach((obj, index) => {
        currentState[index] = { position: obj.position.clone(), rotation: obj.rotation.clone(), scale: obj.scale.x };
    });
    redoStack.push(currentState);
    const prevState = undoStack.pop();
    applyState(prevState);
}
function redo() {
    if (redoStack.length === 0) return;
    const currentState = {};
    loadedObjectsMap.forEach((obj, index) => {
        currentState[index] = { position: obj.position.clone(), rotation: obj.rotation.clone(), scale: obj.scale.x };
    });
    undoStack.push(currentState);
    const nextState = redoStack.pop();
    applyState(nextState);
}
function applyState(state) {
    for (const index in state) {
        const obj = loadedObjectsMap.get(parseInt(index));
        if (obj) {
            obj.position.copy(state[index].position);
            obj.rotation.copy(state[index].rotation);
            obj.scale.setScalar(state[index].scale);
            obj.userData.originalPosition = obj.position.clone();
        }
    }
    buildEditorPanel();
    if (selectedEditorObject) selectEditorObject(selectedEditorObject);
}

// ==================== GRID MODE ====================
// Uses ONE shared offscreen WebGL renderer, copies frames to 2D canvases
// gridDisplayMode: 'model' = 3D GLB viewers (default), 'image' = thumbnail images from N column
let gridModeActive = false;
let gridTransitioning = false;
let gridSharedRenderer = null;
let gridCells = []; // { scene, camera, mesh, canvas2d, ctx, needsRender, rotating, isImage }
let gridAnimId = null;
let gridDisplayMode ='image'; // 'model';  'model' or 'image' — change at runtime via window.setGridDisplayMode('image') KANKA

// Expose for runtime switching: window.setGridDisplayMode('image') or window.setGridDisplayMode('model')
window.setGridDisplayMode = function(mode) {
    if (mode !== 'image' && mode !== 'model') return;
    gridDisplayMode = mode;
    if (gridModeActive) {
        buildGridCells();
    }
};
window.getGridDisplayMode = function() { return gridDisplayMode; };

function getClickableItemObjects() {
    return clickableObjects.filter(o => o.userData.clickable);
}

window.toggleGridMode = function() {
    if (gridTransitioning) return;
    if (gridModeActive) {
        exitGridMode();
    } else {
        enterGridMode();
    }
};

window.switchToGridMode = function() {
    if (gridModeActive || gridTransitioning) return;
    enterGridMode();
};

window.switchTo3DMode = function() {
    if (!gridModeActive || gridTransitioning) return;
    exitGridMode();
};

function enterGridMode() {
    if (gridModeActive || gridTransitioning) return;
    gridTransitioning = true;
    // Activate grid-3d-toggle (slide thumb to right = grid)
    const g3d = document.getElementById('grid-3d-toggle');
    if (g3d) {
        g3d.classList.add('active');
        const track = g3d.querySelector('.grid-3d-track');
        const thumb = g3d.querySelector('.grid-3d-thumb');
        if (track && thumb) {
            thumb.style.left = track.dataset.thumbActiveLeft || '0px';
        }
    }
    // Hide camera dots (row2) on desktop; hide wire toggle (row1) if image mode (both desktop & mobile)
    if (!isMobileView()) {
        const row2 = document.getElementById('right-toggle-row2');
        if (row2) row2.style.display = 'none';
    }
    if (gridDisplayMode === 'image') {
        const row1 = document.getElementById('right-toggle-row1');
        if (row1) row1.style.display = 'none';
    }

    // Close any open panels
    if (document.getElementById('side-panel').classList.contains('open')) closePanel();

    // Pause main scene rendering to save GPU
    sceneDirty = false;

    // Fade to white (or purple in wire mode)
    const fade = document.getElementById('grid-fade');
    if (isWireMode) fade.classList.add('wire-fade');
    else fade.classList.remove('wire-fade');
    fade.classList.add('active');

    setTimeout(() => {
        // Hide main scene
        const sceneContainer = document.getElementById('scene-container');
        sceneContainer.style.visibility = 'hidden';

        // Disable controls
        controls.enabled = false;
        controlsOrtho.enabled = false;

        // Show grid overlay
        const overlay = document.getElementById('grid-overlay');
        if (isWireMode) overlay.classList.add('wire-bg');
        else overlay.classList.remove('wire-bg');
        overlay.classList.add('active');

        // Build grid
        buildGridCells();

        // Fade out the white overlay to reveal grid
        setTimeout(() => {
            fade.classList.remove('active');
            gridModeActive = true;
            gridTransitioning = false;
        }, 100);
    }, 650);
}

function exitGridMode() {
    if (!gridModeActive || gridTransitioning) return;
    gridTransitioning = true;
    // Close grid modal if open
    if (document.getElementById('grid-modal-overlay').classList.contains('open')) {
        window.closeGridModal();
    }
    // Deactivate grid-3d-toggle (slide thumb to left = 3D)
    const g3d = document.getElementById('grid-3d-toggle');
    if (g3d) {
        g3d.classList.remove('active');
        const thumb = g3d.querySelector('.grid-3d-thumb');
        if (thumb) thumb.style.left = '0px';
    }
    // Show camera dots (row2) on desktop; always restore wire toggle (row1) when back to 3D
    if (!isMobileView()) {
        const row2 = document.getElementById('right-toggle-row2');
        if (row2) row2.style.display = 'flex';
    }
    const row1 = document.getElementById('right-toggle-row1');
    if (row1) row1.style.display = 'flex';

    // Fade to white
    const fade = document.getElementById('grid-fade');
    if (isWireMode) fade.classList.add('wire-fade');
    else fade.classList.remove('wire-fade');
    fade.classList.add('active');

    setTimeout(() => {
        // Dispose grid
        disposeGridCells();
        document.getElementById('grid-overlay').classList.remove('active', 'wire-bg');
        document.getElementById('grid-container').innerHTML = '';

        // Show main scene
        const sceneContainer = document.getElementById('scene-container');
        sceneContainer.style.visibility = 'visible';

        // Re-enable controls
        if (currentCameraMode === 'persp') {
            controls.enabled = true;
        } else {
            controlsOrtho.enabled = true;
        }
        markDirty();

        // Fade out white to reveal scene
        setTimeout(() => {
            fade.classList.remove('active');
            gridModeActive = false;
            gridTransitioning = false;
        }, 100);
    }, 650);
}

let gridPendingClones = []; // queue for staggered cloning
// Visibility tracking: only render cells that are visible on screen
let gridVisibleSet = new Set();
let gridIOObserver = null;

function buildGridCells() {
    disposeGridCells();
    gridPendingClones = [];
    const container = document.getElementById('grid-container');
    container.innerHTML = '';

    const isMobileNow = isMobileView();
    const useImageMode = gridDisplayMode === 'image';

    // Sort by Excel row order (rowIndex stored in userData)
    const items = getClickableItemObjects().slice().sort((a, b) => {
        const ai = a.userData.rowIndex !== undefined ? a.userData.rowIndex : a.userData.sceneObjectIndex;
        const bi = b.userData.rowIndex !== undefined ? b.userData.rowIndex : b.userData.sceneObjectIndex;
        return ai - bi;
    });

    // Only create shared renderer for model mode
    if (!useImageMode) {
        const offCanvas = document.createElement('canvas');
        const cellPx = isMobileNow ? 224 : 256;
        offCanvas.width = cellPx;
        offCanvas.height = cellPx;
        gridSharedRenderer = new THREE.WebGLRenderer({ canvas: offCanvas, antialias: true, alpha: true });
        gridSharedRenderer.setSize(cellPx, cellPx, false);
        gridSharedRenderer.setPixelRatio(1);
        gridSharedRenderer.outputColorSpace = THREE.SRGBColorSpace;
        if (!isWireMode) {
            gridSharedRenderer.toneMapping = THREE.ACESFilmicToneMapping;
            gridSharedRenderer.toneMappingExposure = 1.1;
        } else {
            gridSharedRenderer.toneMapping = THREE.NoToneMapping;
        }
        gridSharedRenderer.setClearColor(0x000000, 0);
        gridSharedRenderer.transmissionSampler = true;
    }

    // IntersectionObserver for visibility-based rendering (model mode) or lazy-loading (image mode)
    if (gridIOObserver) { gridIOObserver.disconnect(); gridIOObserver = null; }
    gridVisibleSet = new Set();

    gridIOObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            const idx = parseInt(entry.target.dataset.gridIdx);
            if (isNaN(idx)) return;
            if (entry.isIntersecting) {
                gridVisibleSet.add(idx);
                // For image mode: lazy load the image
                if (useImageMode) {
                    const img = entry.target.querySelector('img[data-src]');
                    if (img && !img.src) {
                        img.src = img.dataset.src;
                        img.removeAttribute('data-src');
                    }
                }
            } else {
                gridVisibleSet.delete(idx);
            }
        });
    }, { root: document.getElementById('grid-overlay'), rootMargin: '100px' });

    items.forEach((mainObj, i) => {
        const cell = document.createElement('div');
        cell.className = 'grid-cell';
        cell.dataset.gridIdx = i;

        const canvasWrap = document.createElement('div');
        canvasWrap.className = 'grid-cell-canvas-wrap';

        const mc = mainObj.userData.modalContent || {};
        const cellPx = isMobileNow ? 224 : 256;

        if (useImageMode) {
            // IMAGE MODE: square cell, image scaled by thumbScale (column AA/26), pinned to bottom-center
            const thumbUrl = mainObj.userData.gridThumbnail || '';
            const thumbScale = mainObj.userData.thumbScale || 1;
            // Container: square, flex-end to pin image to bottom, overflow hidden
            canvasWrap.style.cssText = 'width:100%; aspect-ratio:1/1; position:relative; display:flex; align-items:flex-end; justify-content:center; overflow:hidden;';
            const img = document.createElement('img');
            // Image fills 100% of the box, then transform:scale() visually scales it
            // transform-origin:center bottom keeps bottom-center anchored
            // Box size NEVER changes — zero layout shift
            img.style.cssText = `display:block; width:100%; height:100%; object-fit:contain; object-position:bottom center; transform:scale(${thumbScale}); transform-origin:center bottom;`;
            img.alt = mc.name || '';
            // Lazy load via IntersectionObserver
            img.dataset.src = thumbUrl;
            canvasWrap.appendChild(img);

            // Tap to open grid modal
            canvasWrap.addEventListener('click', () => {
                if (mainObj.userData.modalContent) window.openGridModal(mainObj.userData.modalContent);
            });

            const cellData = {
                scene: null, camera: null, mesh: null,
                canvas2d: null, ctx: null, needsRender: false, rotating: false, mainObj, loaded: true, isImage: true
            };
            gridCells.push(cellData);
        } else {
            // MODEL MODE: 3D canvas rendering (original behavior)
            const canvas2d = document.createElement('canvas');
            canvas2d.width = cellPx;
            canvas2d.height = cellPx;
            const ctx = canvas2d.getContext('2d');
            canvasWrap.appendChild(canvas2d);

            const cellData = {
                scene: null, camera: null, mesh: null,
                canvas2d, ctx, needsRender: false, rotating: false, mainObj, loaded: false, isImage: false
            };
            gridCells.push(cellData);

            // Queue deferred clone (staggered to prevent memory spike)
            gridPendingClones.push({ cellData, mainObj, canvasWrap, cellIndex: i });

            // Touch: distinguish drag (rotate) from tap (popup)
            // On mobile: only horizontal movement rotates; vertical scrolls the grid
            let touchActive = false;
            let touchLastX = 0;
            let touchStartX = 0;
            let touchStartY = 0;
            let touchMoved = false;
            let touchIsHorizontal = null; // null = undecided, true = horizontal drag, false = vertical scroll

            canvasWrap.addEventListener('touchstart', (e) => {
                if (e.touches.length === 1) {
                    touchActive = true;
                    touchLastX = e.touches[0].clientX;
                    touchStartX = e.touches[0].clientX;
                    touchStartY = e.touches[0].clientY;
                    touchMoved = false;
                    touchIsHorizontal = null;
                }
            }, { passive: true });
            canvasWrap.addEventListener('touchmove', (e) => {
                if (!touchActive || e.touches.length !== 1) return;
                const dx = Math.abs(e.touches[0].clientX - touchStartX);
                const dy = Math.abs(e.touches[0].clientY - touchStartY);

                // Decide direction after a small threshold
                if (touchIsHorizontal === null && (dx > 8 || dy > 8)) {
                    touchIsHorizontal = dx > dy;
                }

                if (touchIsHorizontal === true && cellData.mesh) {
                    // Horizontal: rotate model
                    const ddx = e.touches[0].clientX - touchLastX;
                    cellData.mesh.rotation.y += ddx * 0.01;
                    touchLastX = e.touches[0].clientX;
                    cellData.needsRender = true;
                    cellData.rotating = true;
                    touchMoved = true;
                    e.preventDefault();
                } else if (touchIsHorizontal === false) {
                    // Vertical: let browser scroll the grid overlay (don't prevent default)
                    touchMoved = true;
                }
            }, { passive: false });
            canvasWrap.addEventListener('touchend', (e) => {
                const wasDrag = touchMoved;
                touchActive = false;
                cellData.rotating = false;
                touchIsHorizontal = null;
                if (!wasDrag && mainObj.userData.modalContent) {
                    window.openGridModal(mainObj.userData.modalContent);
                }
            });
            canvasWrap.addEventListener('touchcancel', () => { touchActive = false; cellData.rotating = false; touchIsHorizontal = null; });

            // Mouse: drag=rotate, click=popup
            let mDown = false, mLastX = 0, mStartX = 0, mStartY = 0, mMoved = false;
            canvasWrap.addEventListener('mousedown', (e) => {
                mDown = true; mLastX = e.clientX; mStartX = e.clientX; mStartY = e.clientY; mMoved = false;
                cellData.rotating = true;
            });
            window.addEventListener('mousemove', (e) => {
                if (!mDown || !cellData.mesh) return;
                if (Math.abs(e.clientX - mStartX) > 4 || Math.abs(e.clientY - mStartY) > 4) mMoved = true;
                cellData.mesh.rotation.y += (e.clientX - mLastX) * 0.01;
                mLastX = e.clientX;
                cellData.needsRender = true;
            });
            window.addEventListener('mouseup', (e) => {
                if (!mDown) return;
                mDown = false;
                cellData.rotating = false;
                if (!mMoved && mainObj.userData.modalContent) {
                    window.openGridModal(mainObj.userData.modalContent);
                }
            });
        }

        // Label: artist + name
        const label = document.createElement('div');
        label.className = 'grid-cell-label';
        if (isWireMode) label.classList.add('wire-label');
        let labelHtml = '<div class="grid-artist">' + (mc.artist || '') + '</div>';
        labelHtml += '<div class="grid-name">' + (mc.hoverDisplayName || mc.name || '') + '</div>';
        const metaParts = [];
        if (mc.dimensions) metaParts.push(mc.dimensions);
        if (mc.type) metaParts.push(mc.type);
        if (metaParts.length > 0) {
            labelHtml += '<div class="grid-meta">' + metaParts.join(' · ') + '</div>';
        }
        label.innerHTML = labelHtml;
        label.addEventListener('click', () => {
            if (mainObj.userData.modalContent) window.openGridModal(mainObj.userData.modalContent);
        });

        cell.appendChild(canvasWrap);
        cell.appendChild(label);
        container.appendChild(cell);

        // Observe for visibility
        gridIOObserver.observe(cell);
    });

    // Render loop: only for model mode
    if (!useImageMode) {
        let rrIndex = 0;
        let cloneThrottleCounter = 0;
        function gridRenderLoop() {
            gridAnimId = requestAnimationFrame(gridRenderLoop);

            // Process clone queue: 1 model per frame on mobile, 1 on desktop too to limit spikes
            if (gridPendingClones.length > 0) {
                // On mobile, only clone if the cell is visible or near-visible
                const job = gridPendingClones[0];
                const shouldClone = !isMobileNow || gridVisibleSet.has(job.cellIndex) || gridPendingClones.length <= 2;
                if (shouldClone) {
                    gridPendingClones.shift();
                    initGridCellModel(job.cellData, job.mainObj);
                } else {
                    // Move to back of queue
                    cloneThrottleCounter++;
                    if (cloneThrottleCounter > 3) {
                        cloneThrottleCounter = 0;
                        const j = gridPendingClones.shift();
                        initGridCellModel(j.cellData, j.mainObj);
                    }
                }
            }

            if (gridCells.length === 0) return;
            // Only render visible cells, max 2 per frame (1 on mobile)
            const maxPerFrame = isMobileNow ? 1 : 2;
            let rendered = 0;
            for (let attempts = 0; attempts < gridCells.length && rendered < maxPerFrame; attempts++) {
                const cellIdx = rrIndex % gridCells.length;
                const cell = gridCells[cellIdx];
                rrIndex++;
                if (!cell.loaded || cell.isImage) continue;
                // Skip cells that are not visible on screen
                if (!gridVisibleSet.has(cellIdx) && !cell.rotating) continue;
                if (cell.needsRender || cell.rotating) {
                    gridSharedRenderer.render(cell.scene, cell.camera);
                    cell.ctx.clearRect(0, 0, cell.canvas2d.width, cell.canvas2d.height);
                    cell.ctx.drawImage(gridSharedRenderer.domElement, 0, 0);
                    cell.needsRender = false;
                    rendered++;
                }
            }
        }
        gridRenderLoop();
    }
}

function initGridCellModel(cellData, mainObj) {
    const miniScene = new THREE.Scene();
    if (!isWireMode && globalEnvMap) {
        miniScene.environment = globalEnvMap;
    }
    // Studio lighting: bright enough for PBR but env map does the heavy lifting
    const isMobileNow = isMobileView();
    const hemiLight = new THREE.HemisphereLight(0xffffff, 0xd4c5a9, 1.5);
    miniScene.add(hemiLight);
    const keyLight = new THREE.DirectionalLight(0xffffff, 2.0);
    keyLight.position.set(3, 8, 10);
    miniScene.add(keyLight);
    if (!isMobileNow) {
        // Extra lights only on desktop to save GPU on mobile
        const fillLight = new THREE.DirectionalLight(0xddeeff, 0.8);
        fillLight.position.set(-5, 4, -3);
        miniScene.add(fillLight);
        const rimLight = new THREE.DirectionalLight(0xffffff, 0.5);
        rimLight.position.set(0, 2, -8);
        miniScene.add(rimLight);
    }

    // Clone from the ORIGINAL materials stored in originalMaterials map
    // (main scene may have wireframe applied, so we need the real ones)
    const cloned = mainObj.clone(true);
    cloned.position.set(0, 0, 0);
    cloned.rotation.set(0, 0, 0);
    cloned.scale.setScalar(1);

    if (isWireMode) {
        cloned.traverse(child => {
            if (child.isMesh && child.material) {
                child.material = new THREE.MeshBasicMaterial({
                    color: 0xffffff, wireframe: true,
                    transparent: true, opacity: 0.9, side: THREE.DoubleSide
                });
            }
        });
    } else {
        // Build UUID map from original mesh → original material
        // (clone changes UUIDs, so match by traversal order from source)
        const srcMeshes = [];
        mainObj.traverse(child => { if (child.isMesh) srcMeshes.push(child); });
        const clonedMeshes = [];
        cloned.traverse(child => { if (child.isMesh) clonedMeshes.push(child); });

        for (let mi = 0; mi < clonedMeshes.length; mi++) {
            const clonedChild = clonedMeshes[mi];
            const srcChild = srcMeshes[mi];
            if (!srcChild) continue;

            // Get original PBR material (before any wireframe override)
            const origMat = originalMaterials.get(srcChild.uuid) || srcChild.material;
            if (isMobileNow) {
                // On mobile: use simpler material clone to save RAM
                clonedChild.material = origMat.clone();
                clonedChild.material.envMap = null; // skip per-mesh envmap on mobile to save GPU
            } else {
                clonedChild.material = origMat.clone();
            }

            const mat = clonedChild.material;
            // Env map for reflections (desktop only per-mesh)
            if (!isMobileNow && globalEnvMap && (mat.isMeshStandardMaterial || mat.isMeshPhysicalMaterial)) {
                mat.envMap = globalEnvMap;
                mat.envMapIntensity = mat.envMapIntensity || 1.0;
            }
            mat.needsUpdate = true;
            clonedChild.castShadow = false;
            clonedChild.receiveShadow = false;
        }
    }

    // Center X/Z, keep Y at origin (bottom center)
    const box = new THREE.Box3().setFromObject(cloned);
    const size = box.getSize(new THREE.Vector3());
    const center = box.getCenter(new THREE.Vector3());
    cloned.position.set(-center.x, 0, -center.z);

    // AY=scale, AZ=rotX, BA=rotY, BB=height offset
    const ayScale = mainObj.userData.gridScale || 0;
    const azRotX = (mainObj.userData.gridRotX || 0) * Math.PI / 180;
    const baRotY = (mainObj.userData.gridRotY || 0) * Math.PI / 180;
    const bbOffsetY = mainObj.userData.gridOffsetY || 0;

    if (ayScale > 0) {
        cloned.scale.setScalar(ayScale);
    } else {
        const maxDim = Math.max(size.x, size.y, size.z);
        if (maxDim > 0) cloned.scale.setScalar(1.6 / maxDim);
    }
    cloned.rotation.x = azRotX;
    cloned.rotation.y = baRotY;
    cloned.position.y = bbOffsetY;

    miniScene.add(cloned);

    const viewSize = 1;
    const miniCam = new THREE.OrthographicCamera(-viewSize, viewSize, viewSize, -viewSize, 0.01, 1000);
    miniCam.position.set(0, 0, 10);
    miniCam.lookAt(0, 0, 0);

    cellData.scene = miniScene;
    cellData.camera = miniCam;
    cellData.mesh = cloned;
    cellData.needsRender = true;
    cellData.loaded = true;
}

function disposeGridCells() {
    if (gridAnimId) { cancelAnimationFrame(gridAnimId); gridAnimId = null; }
    if (gridIOObserver) { gridIOObserver.disconnect(); gridIOObserver = null; }
    gridVisibleSet = new Set();
    gridPendingClones = [];
    gridCells.forEach(cell => {
        if (cell.isImage) return; // image cells have no Three.js resources
        if (!cell.scene) return;
        cell.scene.traverse(child => {
            if (child.geometry) child.geometry.dispose();
            if (child.material) {
                if (Array.isArray(child.material)) child.material.forEach(m => m.dispose());
                else child.material.dispose();
            }
        });
        cell.scene = null;
        cell.camera = null;
        cell.mesh = null;
    });
    gridCells = [];
    if (gridSharedRenderer) { gridSharedRenderer.dispose(); gridSharedRenderer = null; }
}

// Update wireframe in grid mode when toggled
const _origToggleWireMode = window.toggleWireMode;
window.toggleWireMode = function() {
    _origToggleWireMode();
    if (gridModeActive) {
        const overlay = document.getElementById('grid-overlay');
        if (isWireMode) overlay.classList.add('wire-bg');
        else overlay.classList.remove('wire-bg');
        buildGridCells();
    }
};
