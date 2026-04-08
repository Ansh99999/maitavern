const STORAGE_KEY = "maitavern-state-v8"; // legacy fallback (migration only)
const USER_FILE_STORAGE_KEY = "maitavern-user-file-v1"; // legacy fallback (migration only)
const USER_FILE_SCHEMA_VERSION = 1;
const USER_FILE_NAME = "user";
const USER_FILE_ENDPOINT = "/api/user";

let userFileWriteTimer = null;
let pendingUserFilePayload = null;
let userFileWriteInFlight = false;
let deviceStorageUnavailable = false;

const defaults = {
  hasEnteredApp: false,
  activeView: "homeView",
  endpoint: "",
  apiKey: "",
  lastGenerationType: "",
  authHeader: "Authorization",
  authPrefix: "Bearer ",
  extraHeaders: "",
  extraBody: "",
  models: ["gpt-4o-mini"],
  selectedModel: "gpt-4o-mini",
  systemPrompt: "You are a helpful assistant.",
  messages: [
    {
      role: "system",
      content: "Welcome to MaiTavern. Open settings to add your endpoint, API key, and models.",
    },
  ],
  characters: [],
  providers: [],
  assets: [],
  lorebooks: [],
  lorebookSettings: {
    scanDepth: 4,
    tokenBudget: 1200,
    recursiveScanning: true,
  },
  presets: [
    {
      id: "default-preset",
      name: "Default preset",
      systemPrompt: "You are a helpful assistant.",
      temperature: 0.7,
      topP: 1,
      maxTokens: 300,
      stream: false,
      contextTemplate: "{{#if system}}{{system}}\n{{/if}}{{#if description}}{{description}}\n{{/if}}{{#if personality}}{{personality}}\n{{/if}}{{#if scenario}}{{scenario}}\n{{/if}}{{#if persona}}{{persona}}\n{{/if}}{{trim}}",
      chatStart: "***",
      exampleSeparator: "***",
      instruct: {
        inputSequence: "<|im_start|>user",
        outputSequence: "<|im_start|>assistant",
        systemSequence: "<|im_start|>system",
        stopSequence: "<|im_end|>",
      },
      useSystemPrompt: true,
      namesBehavior: "",
      prompts: [],
      promptOrder: [],
      source: "maitavern",
    },
  ],
  activePresetId: "default-preset",
  activeProviderId: "",
  chats: [],
  recentChats: [],
  currentChatId: "",
  currentChatTitle: "New chat",
  macroGlobalVars: {},
  macroChatVars: {},
  customization: {
    avatarSize: 40,
    defaultTextColor: "#f5f7fb",
    appFontSource: "system",
    appFontFamily: "Inter",
    appFontFileName: "",
    appFontDataUrl: "",
    appFontGoogleName: "",
    appFontGoogleCss: "",
    wrapRules: [
      { id: "quotes-double", title: '"..."', start: '"', end: '"', color: "#7dd3fc", enabled: true },
      { id: "asterisk", title: "*...*", start: "*", end: "*", color: "#f9a8d4", enabled: true },
      { id: "guillemets", title: "«...»", start: "«", end: "»", color: "#c084fc", enabled: true },
    ],
  },
};

let state = structuredClone(defaults);
let currentController = null;
let presetEditorDraft = null;
let messageActionState = { index: -1, mode: "" };
let wrapRuleSearchQuery = "";
let appFontStyleTag = null;
let lorebookEntryDraft = [];
let lorebookEntryLongPress = {
  timer: null,
  pointerId: null,
  sourceIndex: -1,
  targetIndex: -1,
  active: false,
  moveHandler: null,
  upHandler: null,
  autoScrollRaf: null,
  latestClientX: 0,
  latestClientY: 0,
  scrollContainer: null,
};
let lorebookEntrySearchQuery = "";

const els = {
  welcomeScreen: document.getElementById("welcomeScreen"),
  appShell: document.getElementById("appShell"),
  enterAppBtn: document.getElementById("enterAppBtn"),
  homeBrandBtn: document.getElementById("homeBrandBtn"),
  settingsToggle: document.getElementById("settingsToggle"),
  settingsPanel: document.getElementById("settingsPanel"),
  closeSettingsBtn: document.getElementById("closeSettingsBtn"),
  endpointInput: document.getElementById("endpointInput"),
  apiKeyInput: document.getElementById("apiKeyInput"),
  authHeaderInput: document.getElementById("authHeaderInput"),
  authPrefixInput: document.getElementById("authPrefixInput"),
  modelsInput: document.getElementById("modelsInput"),
  extraHeadersInput: document.getElementById("extraHeadersInput"),
  extraBodyInput: document.getElementById("extraBodyInput"),
  systemPromptInput: document.getElementById("systemPromptInput"),
  saveSettingsBtn: document.getElementById("saveSettingsBtn"),
  clearStorageBtn: document.getElementById("clearStorageBtn"),
  fetchModelsBtn: document.getElementById("fetchModelsBtn"),
  fetchModelsStatus: document.getElementById("fetchModelsStatus"),
  modelSelect: document.getElementById("modelSelect"),
  messages: document.getElementById("messages"),
  composer: document.getElementById("composer"),
  messageInput: document.getElementById("messageInput"),
  sendBtn: document.getElementById("sendBtn"),
  stopBtn: document.getElementById("stopBtn"),
  statusText: document.getElementById("statusText"),
  newChatBtn: document.getElementById("newChatBtn"),
  messageTemplate: document.getElementById("messageTemplate"),
  navCards: Array.from(document.querySelectorAll(".nav-card")),
  views: Array.from(document.querySelectorAll(".view")),
  characterList: document.getElementById("characterList"),
  characterSearchInput: document.getElementById("characterSearchInput"),
  characterStudioBtn: document.getElementById("characterStudioBtn"),
  characterStudioImportJsonBtn: document.getElementById("characterStudioImportJsonBtn"),
  characterStudioImportAvatarBtn: document.getElementById("characterStudioImportAvatarBtn"),
  characterImportInput: document.getElementById("characterImportInput"),
  characterAvatarImportInput: document.getElementById("characterAvatarImportInput"),
  characterEditorCard: document.getElementById("characterEditorCard"),
  characterNameInput: document.getElementById("characterNameInput"),
  characterScenarioInput: document.getElementById("characterScenarioInput"),
  characterPersonalityInput: document.getElementById("characterPersonalityInput"),
  characterGreetingInput: document.getElementById("characterGreetingInput"),
  characterAlternateGreetingsInput: document.getElementById("characterAlternateGreetingsInput"),
  characterAvatarInput: document.getElementById("characterAvatarInput"),
  characterAvatarPreview: document.getElementById("characterAvatarPreview"),
  characterSaveBtn: document.getElementById("characterSaveBtn"),
  characterCancelBtn: document.getElementById("characterCancelBtn"),
  lorebookScanDepthInput: document.getElementById("lorebookScanDepthInput"),
  lorebookTokenBudgetInput: document.getElementById("lorebookTokenBudgetInput"),
  lorebookRecursiveInput: document.getElementById("lorebookRecursiveInput"),
  saveLorebookSettingsBtn: document.getElementById("saveLorebookSettingsBtn"),
  lorebookSettingsStatus: document.getElementById("lorebookSettingsStatus"),
  lorebookSearchInput: document.getElementById("lorebookSearchInput"),
  lorebookImportBtn: document.getElementById("lorebookImportBtn"),
  lorebookExportBtn: document.getElementById("lorebookExportBtn"),
  lorebookNewBtn: document.getElementById("lorebookNewBtn"),
  lorebookImportInput: document.getElementById("lorebookImportInput"),
  libraryCategoryChips: Array.from(document.querySelectorAll(".library-category-chip")),
  libraryPanels: Array.from(document.querySelectorAll("#libraryView .library-panel")),
  lorebookEditorView: document.getElementById("lorebookEditorView"),
  lorebookEditorBackBtn: document.getElementById("lorebookEditorBackBtn"),
  lorebookEditorImportBtn: document.getElementById("lorebookEditorImportBtn"),
  lorebookEditorImportInput: document.getElementById("lorebookEditorImportInput"),
  lorebookEditorStatus: document.getElementById("lorebookEditorStatus"),
  lorebookNameInput: document.getElementById("lorebookNameInput"),
  lorebookDescriptionInput: document.getElementById("lorebookDescriptionInput"),
  lorebookSaveBtn: document.getElementById("lorebookSaveBtn"),
  lorebookCancelBtn: document.getElementById("lorebookCancelBtn"),
  lorebookList: document.getElementById("lorebookList"),
  lorebookEntrySearchInput: document.getElementById("lorebookEntrySearchInput"),
  lorebookEntryNewBtn: document.getElementById("lorebookEntryNewBtn"),
  lorebookEntryList: document.getElementById("lorebookEntryList"),
  lorebookEntryEditorCard: document.getElementById("lorebookEntryEditorCard"),
  lorebookEntryTitleInput: document.getElementById("lorebookEntryTitleInput"),
  lorebookEntryKeysInput: document.getElementById("lorebookEntryKeysInput"),
  lorebookEntrySecondaryKeysInput: document.getElementById("lorebookEntrySecondaryKeysInput"),
  lorebookEntryContentInput: document.getElementById("lorebookEntryContentInput"),
  lorebookEntryPriorityInput: document.getElementById("lorebookEntryPriorityInput"),
  lorebookEntryProbabilityInput: document.getElementById("lorebookEntryProbabilityInput"),
  lorebookEntryEnabledInput: document.getElementById("lorebookEntryEnabledInput"),
  lorebookEntryConstantInput: document.getElementById("lorebookEntryConstantInput"),
  lorebookEntryCaseSensitiveInput: document.getElementById("lorebookEntryCaseSensitiveInput"),
  lorebookEntrySelectiveInput: document.getElementById("lorebookEntrySelectiveInput"),
  lorebookEntrySelectiveLogicInput: document.getElementById("lorebookEntrySelectiveLogicInput"),
  lorebookEntrySaveBtn: document.getElementById("lorebookEntrySaveBtn"),
  lorebookEntryDeleteBtn: document.getElementById("lorebookEntryDeleteBtn"),
  lorebookEntryCancelBtn: document.getElementById("lorebookEntryCancelBtn"),
  providerSearchInput: document.getElementById("providerSearchInput"),
  providerImportBtn: document.getElementById("providerImportBtn"),
  providerExportBtn: document.getElementById("providerExportBtn"),
  providerNewBtn: document.getElementById("providerNewBtn"),
  providerImportInput: document.getElementById("providerImportInput"),
  providerFormCard: document.getElementById("providerFormCard"),
  providerNameInput: document.getElementById("providerNameInput"),
  providerUrlInput: document.getElementById("providerUrlInput"),
  providerApiKeyInput: document.getElementById("providerApiKeyInput"),
  providerThinkStartInput: document.getElementById("providerThinkStartInput"),
  providerThinkEndInput: document.getElementById("providerThinkEndInput"),
  providerShowThinkingInput: document.getElementById("providerShowThinkingInput"),
  providerIncludeThinkingInput: document.getElementById("providerIncludeThinkingInput"),
  providerSaveBtn: document.getElementById("providerSaveBtn"),
  providerCancelBtn: document.getElementById("providerCancelBtn"),
  providerFetchModelsBtn: document.getElementById("providerFetchModelsBtn"),
  providerFetchModelsStatus: document.getElementById("providerFetchModelsStatus"),
  providerList: document.getElementById("providerList"),
  presetNewBtn: document.getElementById("presetNewBtn"),
  presetFormCard: document.getElementById("presetFormCard"),
  presetImportBtn: document.getElementById("presetImportBtn"),
  presetExportBtn: document.getElementById("presetExportBtn"),
  presetImportInput: document.getElementById("presetImportInput"),
  presetNameInput: document.getElementById("presetNameInput"),
  presetTemperatureInput: document.getElementById("presetTemperatureInput"),
  presetTopPInput: document.getElementById("presetTopPInput"),
  presetMaxTokensInput: document.getElementById("presetMaxTokensInput"),
  presetStreamInput: document.getElementById("presetStreamInput"),
  presetSystemPromptInput: document.getElementById("presetSystemPromptInput"),
  presetContextTemplateInput: document.getElementById("presetContextTemplateInput"),
  presetChatStartInput: document.getElementById("presetChatStartInput"),
  presetExampleSeparatorInput: document.getElementById("presetExampleSeparatorInput"),
  presetInstructUserInput: document.getElementById("presetInstructUserInput"),
  presetInstructAssistantInput: document.getElementById("presetInstructAssistantInput"),
  presetInstructSystemInput: document.getElementById("presetInstructSystemInput"),
  presetStopSequenceInput: document.getElementById("presetStopSequenceInput"),
  presetUseSystemPromptInput: document.getElementById("presetUseSystemPromptInput"),
  presetNamesBehaviorInput: document.getElementById("presetNamesBehaviorInput"),
  presetScenarioFormatInput: document.getElementById("presetScenarioFormatInput"),
  presetPersonalityFormatInput: document.getElementById("presetPersonalityFormatInput"),
  presetAddPromptBtn: document.getElementById("presetAddPromptBtn"),
  presetPromptList: document.getElementById("presetPromptList"),
  presetPromptModalBackdrop: document.getElementById("presetPromptModalBackdrop"),
  presetPromptEditor: document.getElementById("presetPromptEditor"),
  presetPromptIdentifierInput: document.getElementById("presetPromptIdentifierInput"),
  presetPromptNameInput: document.getElementById("presetPromptNameInput"),
  presetPromptRoleInput: document.getElementById("presetPromptRoleInput"),
  presetPromptInjectionPositionInput: document.getElementById("presetPromptInjectionPositionInput"),
  presetPromptInjectionDepthInput: document.getElementById("presetPromptInjectionDepthInput"),
  presetPromptEnabledInput: document.getElementById("presetPromptEnabledInput"),
  presetPromptSystemPromptInput: document.getElementById("presetPromptSystemPromptInput"),
  presetPromptMarkerInput: document.getElementById("presetPromptMarkerInput"),
  presetPromptForbidOverridesInput: document.getElementById("presetPromptForbidOverridesInput"),
  presetPromptContentInput: document.getElementById("presetPromptContentInput"),
  presetPromptSaveBtn: document.getElementById("presetPromptSaveBtn"),
  presetPromptDeleteBtn: document.getElementById("presetPromptDeleteBtn"),
  presetPromptCloseBtn: document.getElementById("presetPromptCloseBtn"),
  presetSaveBtn: document.getElementById("presetSaveBtn"),
  presetCancelBtn: document.getElementById("presetCancelBtn"),
  presetList: document.getElementById("presetList"),
  recentChats: document.getElementById("recentChats"),
  chatEmptyState: document.getElementById("chatEmptyState"),
  impersonateToggleBtn: document.getElementById("impersonateToggleBtn"),
  impersonateSelect: document.getElementById("impersonateSelect"),
  chatDrawer: document.getElementById("chatDrawer"),
  closeChatDrawerBtn: document.getElementById("closeChatDrawerBtn"),
  chatTitleInput: document.getElementById("chatTitleInput"),
  drawerLinks: Array.from(document.querySelectorAll("#chatDrawer .drawer-link[data-drawer-view]")),
  applyOtherPresetBtn: document.getElementById("applyOtherPresetBtn"),
  applyOtherProviderBtn: document.getElementById("applyOtherProviderBtn"),
  presetPicker: document.getElementById("presetPicker"),
  providerPicker: document.getElementById("providerPicker"),
  drawerPresetList: document.getElementById("drawerPresetList"),
  drawerProviderList: document.getElementById("drawerProviderList"),
  currentPresetLabel: document.getElementById("currentPresetLabel"),
  currentProviderLabel: document.getElementById("currentProviderLabel"),
  drawerPresetEditor: document.getElementById("drawerPresetEditor"),
  drawerPresetNameInput: document.getElementById("drawerPresetNameInput"),
  drawerPresetTempInput: document.getElementById("drawerPresetTempInput"),
  drawerPresetSystemInput: document.getElementById("drawerPresetSystemInput"),
  drawerPresetSaveBtn: document.getElementById("drawerPresetSaveBtn"),
  drawerPresetCancelBtn: document.getElementById("drawerPresetCancelBtn"),
  drawerProviderEditor: document.getElementById("drawerProviderEditor"),
  drawerProviderNameInput: document.getElementById("drawerProviderNameInput"),
  drawerProviderUrlInput: document.getElementById("drawerProviderUrlInput"),
  drawerProviderApiKeyInput: document.getElementById("drawerProviderApiKeyInput"),
  drawerProviderThinkStartInput: document.getElementById("drawerProviderThinkStartInput"),
  drawerProviderThinkEndInput: document.getElementById("drawerProviderThinkEndInput"),
  drawerProviderShowThinkingInput: document.getElementById("drawerProviderShowThinkingInput"),
  drawerProviderIncludeThinkingInput: document.getElementById("drawerProviderIncludeThinkingInput"),
  drawerProviderSaveBtn: document.getElementById("drawerProviderSaveBtn"),
  drawerProviderCancelBtn: document.getElementById("drawerProviderCancelBtn"),
  drawerLorebookSearchInput: document.getElementById("drawerLorebookSearchInput"),
  drawerLorebookList: document.getElementById("drawerLorebookList"),
  currentLorebookLabel: document.getElementById("currentLorebookLabel"),
  drawerLorebookScanDepthInput: document.getElementById("drawerLorebookScanDepthInput"),
  drawerLorebookTokenBudgetInput: document.getElementById("drawerLorebookTokenBudgetInput"),
  drawerLorebookRecursiveInput: document.getElementById("drawerLorebookRecursiveInput"),
  saveDrawerLorebookSettingsBtn: document.getElementById("saveDrawerLorebookSettingsBtn"),
  drawerLorebookStatus: document.getElementById("drawerLorebookStatus"),
  messageActionPopup: document.getElementById("messageActionPopup"),
  messageActionCard: document.getElementById("messageActionCard"),
  messageActionEditBtn: document.getElementById("messageActionEditBtn"),
  messageActionDeleteBtn: document.getElementById("messageActionDeleteBtn"),
  messageActionRegenerateBtn: document.getElementById("messageActionRegenerateBtn"),
  messageActionForkBtn: document.getElementById("messageActionForkBtn"),
  messageDeleteOptions: document.getElementById("messageDeleteOptions"),
  messageDeletePromptText: document.getElementById("messageDeletePromptText"),
  deleteOneMessageBtn: document.getElementById("deleteOneMessageBtn"),
  deleteFromHereBtn: document.getElementById("deleteFromHereBtn"),
  messageEditFieldWrap: document.getElementById("messageEditFieldWrap"),
  messageEditInput: document.getElementById("messageEditInput"),
  messageActionConfirmBtn: document.getElementById("messageActionConfirmBtn"),
  messageActionCloseBtn: document.getElementById("messageActionCloseBtn"),
  toastContainer: document.getElementById("toastContainer"),
  avatarSizeInput: document.getElementById("avatarSizeInput"),
  defaultTextColorInput: document.getElementById("defaultTextColorInput"),
  saveCustomizationBtn: document.getElementById("saveCustomizationBtn"),
  chatCustomizationSectionBtn: document.getElementById("chatCustomizationSectionBtn"),
  chatCustomizationSection: document.getElementById("chatCustomizationSection"),
  avatarSizeSectionBtn: document.getElementById("avatarSizeSectionBtn"),
  avatarSizeSection: document.getElementById("avatarSizeSection"),
  textWrappingSectionBtn: document.getElementById("textWrappingSectionBtn"),
  textWrappingSection: document.getElementById("textWrappingSection"),
  wrapRuleSearchInput: document.getElementById("wrapRuleSearchInput"),
  wrapRuleList: document.getElementById("wrapRuleList"),
  wrapRuleEditor: document.getElementById("wrapRuleEditor"),
  wrapRuleTitleInput: document.getElementById("wrapRuleTitleInput"),
  wrapRuleStartInput: document.getElementById("wrapRuleStartInput"),
  wrapRuleEndInput: document.getElementById("wrapRuleEndInput"),
  wrapRuleColorInput: document.getElementById("wrapRuleColorInput"),
  wrapRuleHexInput: document.getElementById("wrapRuleHexInput"),
  addWrapRuleBtn: document.getElementById("addWrapRuleBtn"),
  saveWrapRuleBtn: document.getElementById("saveWrapRuleBtn"),
  deleteWrapRuleBtn: document.getElementById("deleteWrapRuleBtn"),
  closeWrapRuleEditorBtn: document.getElementById("closeWrapRuleEditorBtn"),
  textFontSectionBtn: document.getElementById("textFontSectionBtn"),
  textFontSection: document.getElementById("textFontSection"),
  openFontManagerBtn: document.getElementById("openFontManagerBtn"),
  currentFontName: document.getElementById("currentFontName"),
  fontManagerBackdrop: document.getElementById("fontManagerBackdrop"),
  fontManagerCard: document.getElementById("fontManagerCard"),
  closeFontManagerBtn: document.getElementById("closeFontManagerBtn"),
  fontFileInput: document.getElementById("fontFileInput"),
  browseFontFileBtn: document.getElementById("browseFontFileBtn"),
  uploadedFontMeta: document.getElementById("uploadedFontMeta"),
  googleFontSearchInput: document.getElementById("googleFontSearchInput"),
  addGoogleFontBtn: document.getElementById("addGoogleFontBtn"),
  fontManagerStatus: document.getElementById("fontManagerStatus"),
  resetAppFontBtn: document.getElementById("resetAppFontBtn"),
  exportDataBtn: document.getElementById("exportDataBtn"),
  importDataBtn: document.getElementById("importDataBtn"),
  importDataInput: document.getElementById("importDataInput"),
  dataStatus: document.getElementById("dataStatus"),
};

void init();

async function init() {
  state = await loadState();
  ensureCollections();
  syncInputsFromState();
  renderModelOptions();
  applyCustomization();
  renderMessages();
  renderCharacters();
  renderProviders();
  renderPresets();
  renderLorebooks();
  setActiveLibraryPanel("libraryLorebooksPanel");
  renderRecentChats();
  renderImpersonationOptions();
  renderDrawerState();
  syncCustomizationInputs();
  syncLorebookSettingsInputs();
  renderWrapRuleList();
  applyWelcomeState();
  switchView(state.activeView || "homeView");
  bindEvents();
  autoResizeTextarea(els.messageInput);
  updateChatUiVisibility();
  maybeAutoFetchModels();
  persistState();
}

function ensureCollections() {
  state.characters = Array.isArray(state.characters)
    ? state.characters.map((character) => normalizeCharacterRecord(character))
    : [];
  state.providers = Array.isArray(state.providers) ? state.providers : [];
  state.assets = Array.isArray(state.assets) ? state.assets : [];
  state.lorebooks = Array.isArray(state.lorebooks)
    ? state.lorebooks.map((lorebook) => normalizeLorebookRecord(lorebook))
    : [];
  state.lorebookSettings = {
    ...structuredClone(defaults.lorebookSettings),
    ...(state.lorebookSettings && typeof state.lorebookSettings === "object" ? state.lorebookSettings : {}),
  };
  state.lorebookSettings.scanDepth = Math.min(64, Math.max(1, Number.parseInt(state.lorebookSettings.scanDepth, 10) || defaults.lorebookSettings.scanDepth));
  state.lorebookSettings.tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(state.lorebookSettings.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget));
  state.lorebookSettings.recursiveScanning = state.lorebookSettings.recursiveScanning !== false;
  state.presets = Array.isArray(state.presets) && state.presets.length ? state.presets : structuredClone(defaults.presets);
  state.presets = state.presets.map(normalizePreset);
  state.chats = Array.isArray(state.chats)
    ? state.chats.map((chat) => normalizeChatRecord(chat))
    : [];
  state.recentChats = Array.isArray(state.recentChats) ? state.recentChats : [];
  state.currentChatId = typeof state.currentChatId === "string" ? state.currentChatId : "";
  state.currentChatTitle = typeof state.currentChatTitle === "string" && state.currentChatTitle.trim() ? state.currentChatTitle : "New chat";
  state.lastGenerationType = typeof state.lastGenerationType === "string" ? state.lastGenerationType : "";
  state.customization = {
    ...structuredClone(defaults.customization),
    ...(state.customization && typeof state.customization === "object" ? state.customization : {}),
  };
  state.macroGlobalVars = state.macroGlobalVars && typeof state.macroGlobalVars === "object" ? state.macroGlobalVars : {};
  state.macroChatVars = state.macroChatVars && typeof state.macroChatVars === "object" ? state.macroChatVars : {};
  state.customization.wrapRules = Array.isArray(state.customization.wrapRules) && state.customization.wrapRules.length
    ? state.customization.wrapRules
    : structuredClone(defaults.customization.wrapRules);
  state.customization.wrapRules = state.customization.wrapRules.map((rule, index) => ({
    id: typeof rule?.id === "string" && rule.id.trim() ? rule.id : crypto.randomUUID(),
    title: typeof rule?.title === "string" ? rule.title : `Wrap ${index + 1}`,
    start: String(rule?.start ?? ""),
    end: String(rule?.end ?? ""),
    color: isValidHexColor(rule?.color) ? rule.color : "#c084fc",
    enabled: rule?.enabled !== false,
  })).filter((rule) => rule.start && rule.end);
  if (!state.customization.wrapRules.length) {
    state.customization.wrapRules = structuredClone(defaults.customization.wrapRules);
  }
  state.customization.appFontSource = typeof state.customization.appFontSource === "string" ? state.customization.appFontSource : defaults.customization.appFontSource;
  if (!["system", "file", "google"].includes(state.customization.appFontSource)) {
    state.customization.appFontSource = defaults.customization.appFontSource;
  }
  state.customization.appFontFamily = typeof state.customization.appFontFamily === "string" && state.customization.appFontFamily.trim()
    ? state.customization.appFontFamily.trim()
    : defaults.customization.appFontFamily;
  state.customization.appFontFileName = typeof state.customization.appFontFileName === "string" ? state.customization.appFontFileName : "";
  state.customization.appFontDataUrl = typeof state.customization.appFontDataUrl === "string" ? state.customization.appFontDataUrl : "";
  state.customization.appFontGoogleName = typeof state.customization.appFontGoogleName === "string" ? state.customization.appFontGoogleName : "";
  state.customization.appFontGoogleCss = typeof state.customization.appFontGoogleCss === "string" ? state.customization.appFontGoogleCss : "";

  if (state.activeProviderId && !state.providers.some((provider) => provider.id === state.activeProviderId)) {
    state.activeProviderId = "";
  }

  if ((!state.endpoint || !state.apiKey) && state.activeProviderId) {
    const activeProvider = state.providers.find((provider) => provider.id === state.activeProviderId);
    if (activeProvider) {
      if (!state.endpoint) state.endpoint = activeProvider.baseUrl || "";
      if (!state.apiKey) state.apiKey = activeProvider.apiKey || "";
    }
  }

  migrateLegacyChats();
}

function bindEvents() {
  els.enterAppBtn?.addEventListener("click", enterApp);
  els.homeBrandBtn?.addEventListener("click", goHome);
  els.settingsToggle?.addEventListener("click", () => {
    if (state.activeView === "chatsView") {
      toggleChatDrawer(true);
    }
  });
  els.closeSettingsBtn?.addEventListener("click", () => toggleSettings(false));
  els.saveSettingsBtn.addEventListener("click", saveSettings);
  els.clearStorageBtn.addEventListener("click", clearStorage);
  els.fetchModelsBtn?.addEventListener("click", () => fetchModelsIntoSettings());
  els.modelSelect.addEventListener("change", () => {
    state.selectedModel = els.modelSelect.value;
    persistState();
  });
  els.composer.addEventListener("submit", (event) => event.preventDefault());
  els.sendBtn.addEventListener("click", onSend);
  els.stopBtn.addEventListener("click", stopRequest);
  els.newChatBtn.addEventListener("click", newChat);
  els.messageInput.addEventListener("input", () => autoResizeTextarea(els.messageInput));
  els.messageInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
    }
  });
  els.navCards.forEach((card) => {
    card.addEventListener("click", () => switchView(card.dataset.view));
  });
  els.characterSearchInput.addEventListener("input", renderCharacters);
  els.characterStudioBtn.addEventListener("click", createCharacter);
  els.characterStudioImportJsonBtn?.addEventListener("click", () => els.characterImportInput.click());
  els.characterStudioImportAvatarBtn?.addEventListener("click", () => els.characterAvatarImportInput.click());
  els.characterSaveBtn.addEventListener("click", saveCharacterFromEditor);
  els.characterCancelBtn.addEventListener("click", hideCharacterEditor);
  els.characterImportInput.addEventListener("change", importCharacters);
  els.characterAvatarImportInput.addEventListener("change", importCharacterFromAvatar);
  els.characterAvatarInput.addEventListener("change", onCharacterAvatarSelected);
  els.libraryCategoryChips?.forEach((chip) => {
    chip.addEventListener("click", () => setActiveLibraryPanel(chip.dataset.libraryPanel || "libraryLorebooksPanel"));
  });
  els.lorebookSearchInput?.addEventListener("input", renderLorebooks);
  els.lorebookImportBtn?.addEventListener("click", () => els.lorebookImportInput?.click());
  els.lorebookImportInput?.addEventListener("change", importLorebooks);
  els.lorebookExportBtn?.addEventListener("click", exportLorebooks);
  els.lorebookNewBtn?.addEventListener("click", () => showLorebookForm());
  els.lorebookEditorImportBtn?.addEventListener("click", () => els.lorebookEditorImportInput?.click());
  els.lorebookEditorImportInput?.addEventListener("change", importLorebookIntoEditor);
  els.lorebookEditorBackBtn?.addEventListener("click", hideLorebookForm);
  els.lorebookSaveBtn?.addEventListener("click", saveLorebookFromForm);
  els.lorebookCancelBtn?.addEventListener("click", hideLorebookForm);
  els.lorebookEntrySearchInput?.addEventListener("input", () => {
    lorebookEntrySearchQuery = els.lorebookEntrySearchInput.value || "";
    renderLorebookEntryDraft();
  });
  els.lorebookEntryNewBtn?.addEventListener("click", () => openLorebookEntryEditor());
  els.lorebookEntrySaveBtn?.addEventListener("click", saveLorebookEntryFromEditor);
  els.lorebookEntryDeleteBtn?.addEventListener("click", deleteLorebookEntryFromEditor);
  els.lorebookEntryCancelBtn?.addEventListener("click", closeLorebookEntryEditor);
  els.saveLorebookSettingsBtn?.addEventListener("click", saveLorebookSettingsFromInputs);
  els.drawerLorebookSearchInput?.addEventListener("input", renderDrawerLorebooks);
  els.saveDrawerLorebookSettingsBtn?.addEventListener("click", saveDrawerLorebookSettings);
  els.providerSearchInput.addEventListener("input", renderProviders);
  els.providerNewBtn.addEventListener("click", showProviderForm);
  els.providerCancelBtn.addEventListener("click", hideProviderForm);
  els.providerSaveBtn.addEventListener("click", saveProviderFromForm);
  els.providerFetchModelsBtn?.addEventListener("click", () => fetchModelsIntoProviderForm());
  els.providerImportBtn.addEventListener("click", () => els.providerImportInput.click());
  els.providerImportInput.addEventListener("change", importProviders);
  els.providerExportBtn.addEventListener("click", exportProviders);
  els.presetImportBtn?.addEventListener("click", () => els.presetImportInput.click());
  els.presetImportInput?.addEventListener("change", importPresets);
  els.presetExportBtn?.addEventListener("click", exportPresets);
  els.presetNewBtn?.addEventListener("click", () => showPresetForm());
  els.presetCancelBtn?.addEventListener("click", hidePresetForm);
  els.presetSaveBtn?.addEventListener("click", savePresetFromForm);
  els.presetAddPromptBtn?.addEventListener("click", createPresetPromptDraft);
  els.presetPromptSaveBtn?.addEventListener("click", savePresetPromptDraft);
  els.presetPromptDeleteBtn?.addEventListener("click", deletePresetPromptDraft);
  els.presetPromptCloseBtn?.addEventListener("click", hidePresetPromptEditor);
  els.presetPromptModalBackdrop?.addEventListener("click", hidePresetPromptEditor);
  els.closeChatDrawerBtn.addEventListener("click", () => toggleChatDrawer(false));
  els.drawerLinks.forEach((button) => {
    button.addEventListener("click", () => {
      if (!button.dataset.drawerView) return;
      switchView(button.dataset.drawerView);
      toggleChatDrawer(false);
    });
  });
  els.applyOtherPresetBtn.addEventListener("click", () => {
    const shouldShow = els.presetPicker.classList.contains("hidden");
    els.presetPicker.classList.toggle("hidden", !shouldShow);
    els.providerPicker.classList.add("hidden");
    hideDrawerProviderEditor();
  });
  els.applyOtherProviderBtn.addEventListener("click", () => {
    const shouldShow = els.providerPicker.classList.contains("hidden");
    els.providerPicker.classList.toggle("hidden", !shouldShow);
    els.presetPicker.classList.add("hidden");
    hideDrawerPresetEditor();
  });
  els.chatTitleInput.addEventListener("input", onChatTitleInput);
  els.impersonateToggleBtn?.addEventListener("click", toggleImpersonationPicker);
  els.impersonateSelect?.addEventListener("change", onImpersonationChange);
  els.saveCustomizationBtn?.addEventListener("click", saveCustomization);
  els.chatCustomizationSectionBtn?.addEventListener("click", () => toggleSection(els.chatCustomizationSection));
  els.avatarSizeSectionBtn?.addEventListener("click", () => toggleSection(els.avatarSizeSection));
  els.textWrappingSectionBtn?.addEventListener("click", () => toggleSection(els.textWrappingSection));
  els.addWrapRuleBtn?.addEventListener("click", createWrapRule);
  els.saveWrapRuleBtn?.addEventListener("click", saveWrapRuleFromEditor);
  els.deleteWrapRuleBtn?.addEventListener("click", deleteWrapRuleFromEditor);
  els.closeWrapRuleEditorBtn?.addEventListener("click", hideWrapRuleEditor);
  els.textFontSectionBtn?.addEventListener("click", () => toggleSection(els.textFontSection));
  els.openFontManagerBtn?.addEventListener("click", openTextFontManager);
  els.closeFontManagerBtn?.addEventListener("click", () => toggleFontManager(false));
  els.fontManagerBackdrop?.addEventListener("click", (event) => {
    if (event.target === els.fontManagerBackdrop) {
      toggleFontManager(false);
    }
  });
  els.browseFontFileBtn?.addEventListener("click", () => els.fontFileInput?.click());
  els.fontFileInput?.addEventListener("change", onFontFileSelected);
  els.addGoogleFontBtn?.addEventListener("click", addFontFromGoogleFonts);
  els.googleFontSearchInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      addFontFromGoogleFonts();
    }
  });
  els.resetAppFontBtn?.addEventListener("click", resetAppFont);

  els.defaultTextColorInput?.addEventListener("input", onDefaultTextColorLiveChange);
  els.wrapRuleColorInput?.addEventListener("input", onWrapRuleColorInputChange);
  els.wrapRuleHexInput?.addEventListener("input", onWrapRuleHexInputChange);
  els.wrapRuleSearchInput?.addEventListener("input", onWrapRuleSearchInput);

  els.exportDataBtn?.addEventListener("click", exportUserDataZip);
  els.importDataBtn?.addEventListener("click", () => els.importDataInput?.click());
  els.importDataInput?.addEventListener("change", importUserDataZip);

  els.drawerPresetSaveBtn?.addEventListener("click", saveDrawerPresetEdits);
  els.drawerPresetCancelBtn?.addEventListener("click", hideDrawerPresetEditor);
  els.drawerProviderSaveBtn?.addEventListener("click", saveDrawerProviderEdits);
  els.drawerProviderCancelBtn?.addEventListener("click", hideDrawerProviderEditor);

  els.messageActionCloseBtn?.addEventListener("click", closeMessageActionPopup);
  els.messageActionPopup?.addEventListener("click", (event) => {
    if (event.target === els.messageActionPopup) closeMessageActionPopup();
  });
  els.messageActionEditBtn?.addEventListener("click", startMessageEditFlow);
  els.messageActionDeleteBtn?.addEventListener("click", startMessageDeleteFlow);
  els.messageActionRegenerateBtn?.addEventListener("click", regenerateFromMessageAction);
  els.messageActionForkBtn?.addEventListener("click", () => showToast("Fork is coming soon", "success"));
  els.messageActionConfirmBtn?.addEventListener("click", commitMessageEditFlow);
  els.deleteOneMessageBtn?.addEventListener("click", () => commitMessageDeleteFlow(false));
  els.deleteFromHereBtn?.addEventListener("click", () => commitMessageDeleteFlow(true));

  document.addEventListener("click", (event) => {
    const trigger = event.target?.closest?.("#textFontSectionBtn, #openFontManagerBtn");
    if (!trigger) return;
    event.preventDefault();
    openTextFontManager();
  });

  window.addEventListener("resize", updateComposerOffset);
}

function enterApp() {
  state.hasEnteredApp = true;
  persistState();
  applyWelcomeState();
}

function goHome() {
  toggleSettings(false);
  toggleChatDrawer(false);
  switchView("homeView");
  setStatus("Home");
}

function applyWelcomeState() {
  els.welcomeScreen.classList.toggle("hidden", state.hasEnteredApp);
  els.appShell.classList.toggle("hidden", !state.hasEnteredApp);
}

function switchView(viewId) {
  state.activeView = viewId;
  els.views.forEach((view) => view.classList.toggle("active", view.id === viewId));
  els.navCards.forEach((card) => {
    const isDirect = card.dataset.view === viewId;
    const isLorebookSubpage = viewId === "lorebookEditorView" && card.dataset.view === "libraryView";
    card.classList.toggle("active", isDirect || isLorebookSubpage);
  });
  els.settingsToggle?.classList.toggle("hidden", viewId !== "chatsView");
  if (viewId !== "chatsView") toggleChatDrawer(false);
  updateChatUiVisibility();
  persistState();
}

function toggleSettings(show) {
  els.settingsPanel.classList.toggle("hidden", !show);
}

function syncInputsFromState() {
  els.endpointInput.value = state.endpoint;
  els.apiKeyInput.value = state.apiKey;
  els.authHeaderInput.value = state.authHeader;
  els.authPrefixInput.value = state.authPrefix;
  els.modelsInput.value = state.models.join("\n");
  els.extraHeadersInput.value = state.extraHeaders;
  els.extraBodyInput.value = state.extraBody;
  if (els.systemPromptInput) {
    els.systemPromptInput.value = state.systemPrompt;
  }
  syncLorebookSettingsInputs();
  setInlineStatus(els.fetchModelsStatus, "");
  setInlineStatus(els.providerFetchModelsStatus, "");
}

function saveSettings() {
  const models = els.modelsInput.value
    .split(/\r?\n/)
    .map((m) => m.trim())
    .filter(Boolean);

  state.endpoint = els.endpointInput.value.trim();
  state.apiKey = els.apiKeyInput.value.trim();
  state.authHeader = els.authHeaderInput.value.trim() || defaults.authHeader;
  state.authPrefix = els.authPrefixInput.value;
  state.extraHeaders = els.extraHeadersInput.value.trim();
  state.extraBody = els.extraBodyInput.value.trim();
  state.models = models.length ? models : defaults.models;
  state.selectedModel = state.models.includes(state.selectedModel) ? state.selectedModel : state.models[0];
  state.systemPrompt = els.systemPromptInput?.value?.trim() || state.systemPrompt || defaults.systemPrompt;

  try {
    parseJsonObject(state.extraHeaders, "Extra headers");
    parseJsonObject(state.extraBody, "Extra body fields");
  } catch (error) {
    setStatus(error.message);
    return;
  }

  persistState();
  renderModelOptions();
  renderDrawerState();
  setStatus("Settings saved");
  toggleSettings(false);
}

async function clearStorage() {
  state = structuredClone(defaults);
  ensureCollections();
  syncInputsFromState();
  renderMessages();
  renderCharacters();
  renderProviders();
  renderPresets();
  renderLorebooks();
  renderRecentChats();
  renderImpersonationOptions();
  renderDrawerState();
  syncLorebookSettingsInputs();
  applyWelcomeState();
  switchView("homeView");
  maybeAutoFetchModels(true);

  if (hasFileSystemAccessApi()) {
    try {
      await writeUserFileToDisk(createUserFilePayload());
      setStatus(`Saved data cleared (wrote ${USER_FILE_NAME})`);
      showToast(`Cleared and wrote ${USER_FILE_NAME}`, "success");
      return;
    } catch {
      // fallback below
    }
  }

  persistState();
  setStatus("Saved data cleared");
}

function createChatRecord(title = "New chat") {
  const chat = {
    id: crypto.randomUUID(),
    title,
    characterId: "",
    lorebookIds: [],
    lorebookSettings: structuredClone(defaults.lorebookSettings),
    messages: [],
    preview: "",
    updatedAt: new Date().toISOString(),
  };
  state.chats.unshift(chat);
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title;
  return chat;
}

function ensureGreetingMessageForChat(chat) {
  if (!chat || !chat.characterId) return;
  const hasAnyMessage = Array.isArray(chat.messages) && chat.messages.length > 0;
  if (hasAnyMessage) return;

  const character = state.characters.find((item) => item.id === chat.characterId);
  if (!character) return;

  const greetingVariants = getCharacterGreetingVariants(character)
    .map((value) => applyMacrosToText(value, character))
    .map((value) => value.trim())
    .filter(Boolean);

  if (chat.messages?.length && chat.messages[0]?.isGreeting && Array.isArray(chat.messages[0]?.swipes)) {
    chat.messages[0].swipes = greetingVariants;
    chat.messages[0].swipe_id = clampSwipeIndex(chat.messages[0].swipe_id, greetingVariants.length);
    chat.messages[0].content = greetingVariants[chat.messages[0].swipe_id] || greetingVariants[0] || "";
    chat.messages[0].swipeable = greetingVariants.length > 1;
    return;
  }


  if (!greetingVariants.length) return;

  chat.messages = [{
    role: "assistant",
    content: greetingVariants[0],
    swipes: greetingVariants,
    swipe_id: 0,
    swipeable: greetingVariants.length > 1,
    isGreeting: true,
  }];
}

function getCurrentChat() {
  return state.chats.find((chat) => chat.id === state.currentChatId) || null;
}

function ensureCurrentChat() {
  return getCurrentChat() || createChatRecord(state.currentChatTitle || "New chat");
}

function syncCurrentChatFromState() {
  const chat = ensureCurrentChat();
  chat.title = state.currentChatTitle || chat.title || "New chat";
  chat.messages = state.messages.map((message) => ({ ...message }));
  chat.lorebookIds = Array.isArray(chat.lorebookIds) ? [...new Set(chat.lorebookIds.map((id) => String(id || "").trim()).filter(Boolean))] : [];
  chat.lorebookSettings = {
    ...structuredClone(defaults.lorebookSettings),
    ...(chat.lorebookSettings && typeof chat.lorebookSettings === "object" ? chat.lorebookSettings : {}),
  };
  chat.lorebookSettings.scanDepth = Math.min(64, Math.max(1, Number.parseInt(chat.lorebookSettings.scanDepth, 10) || defaults.lorebookSettings.scanDepth));
  chat.lorebookSettings.tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(chat.lorebookSettings.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget));
  chat.lorebookSettings.recursiveScanning = chat.lorebookSettings.recursiveScanning !== false;

  const previewSource = [...chat.messages].reverse().find((message) => message.role === "assistant" || message.role === "user");
  chat.preview = (previewSource?.content || "").slice(0, 80);
  chat.updatedAt = new Date().toISOString();
  state.recentChats = state.chats
    .slice()
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
    .slice(0, 8)
    .map(({ id, title, preview, updatedAt, characterId }) => ({ id, title, preview, updatedAt, characterId }));
}

function newChat() {
  switchView("chatsView");
  state.messages = [];
  state.currentChatId = "";
  state.currentChatTitle = "New chat";
  renderMessages();
  renderRecentChats();
  renderDrawerState();
  updateChatUiVisibility();
  setStatus("Choose an existing chat or start one from a character");
}

async function onSend(event) {
  event?.preventDefault?.();
  const currentCharacter = getCurrentChatCharacter();
  const content = applyMacrosToText(els.messageInput.value.trim(), currentCharacter);
  if (!content) return;

  if (!state.endpoint || !state.apiKey) {
    setStatus("Add endpoint and API key in settings");
    toggleSettings(true);
    return;
  }

  switchView("chatsView");
  ensureCurrentChat();

  const userMessage = { role: "user", content };
  state.messages.push(userMessage);
  state.lastGenerationType = "user";

  els.messageInput.value = "";
  autoResizeTextarea(els.messageInput);

  const assistantMessage = { role: "assistant", content: "", pending: true };
  state.messages.push(assistantMessage);

  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();

  const requestMessages = state.messages.slice(0, -1);
  await generateAssistantReply({ requestMessages, assistantMessage });
}

function extractThinkingSummaryFromText(text = "", provider = null) {
  const source = String(text || "");
  const startTag = provider?.thinkingStartTag || "<thinking>";
  const endTag = provider?.thinkingEndTag || "</thinking>";
  if (!startTag || !endTag) return "";
  const escapedStart = escapeRegex(startTag);
  const escapedEnd = escapeRegex(endTag);
  const regex = new RegExp(`${escapedStart}([\\s\\S]*?)${escapedEnd}`, "gi");
  const chunks = [];
  let match;
  while ((match = regex.exec(source)) !== null) {
    chunks.push(String(match[1] || "").trim());
  }
  return chunks.filter(Boolean).join("\n").trim();
}

function stripThinkingBlocks(text = "", provider = null) {
  const source = String(text || "");
  const startTag = provider?.thinkingStartTag || "<thinking>";
  const endTag = provider?.thinkingEndTag || "</thinking>";
  if (!startTag || !endTag) return source;
  const escapedStart = escapeRegex(startTag);
  const escapedEnd = escapeRegex(endTag);
  return source.replace(new RegExp(`${escapedStart}[\\s\\S]*?${escapedEnd}`, "gi"), "").trim();
}

function getActiveProvider() {
  return state.providers.find((item) => item.id === state.activeProviderId) || null;
}

async function generateAssistantReply({ requestMessages, assistantMessage }) {
  setBusy(true);
  setStatus(`Calling ${state.selectedModel}...`);
  currentController = new AbortController();

  try {
    const extraHeaders = parseJsonObject(state.extraHeaders, "Extra headers");
    const extraBody = parseJsonObject(state.extraBody, "Extra body fields");
    const headers = {
      "Content-Type": "application/json",
      ...extraHeaders,
    };

    if (state.apiKey && state.authHeader) {
      headers[state.authHeader] = `${state.authPrefix || ""}${state.apiKey}`;
    }

    const preset = getActivePreset();
    const basePayload = {
      model: state.selectedModel,
      messages: buildApiMessages(requestMessages),
      stream: Boolean(preset?.stream),
      ...extraBody,
    };

    const standardPayload = {
      ...basePayload,
      temperature: preset.temperature ?? 0.7,
      top_p: preset.topP ?? 1,
      max_tokens: preset.maxTokens ?? 300,
      ...(shouldApplyInstructPreset(preset) && preset.instruct?.stopSequence
        ? { stop: [preset.instruct.stopSequence] }
        : {}),
    };

    let data = await requestChatCompletion({ headers, payload: standardPayload, signal: currentController.signal });
    let text = extractAssistantText(data);

    if (!text && shouldRetryWithCompatPayload(data)) {
      const compatPayload = {
        ...basePayload,
        max_completion_tokens: preset.maxTokens ?? 300,
        ...(shouldApplyInstructPreset(preset) && preset.instruct?.stopSequence
          ? { stop: [preset.instruct.stopSequence] }
          : {}),
      };
      data = await requestChatCompletion({ headers, payload: compatPayload, signal: currentController.signal });
      text = extractAssistantText(data);
    }

    if (!text) {
      console.warn("MaiTavern: empty or unrecognized response payload", data);
    }

    const activeProvider = getActiveProvider();
    const rawText = text || `Provider returned an empty assistant message. ${JSON.stringify(data).slice(0, 280)}`;
    assistantMessage.pending = false;
    assistantMessage.thinkingSummary = extractThinkingSummaryFromText(rawText, activeProvider);
    assistantMessage.content = stripThinkingBlocks(rawText, activeProvider);
    state.lastGenerationType = "assistant";

    syncCurrentChatFromState();
    persistState();
    renderMessages();
    renderRecentChats();
    setStatus(text ? "Done" : "Provider returned empty content");
  } catch (error) {
    assistantMessage.pending = false;
    assistantMessage.content = `Error: ${error.message}`;
    syncCurrentChatFromState();
    persistState();
    renderMessages();
    renderRecentChats();
    setStatus(error.name === "AbortError" ? "Request stopped" : "Request failed");
  } finally {
    setBusy(false);
    currentController = null;
  }
}

function stopRequest() {
  currentController?.abort();
}

function buildApiMessages(messages) {
  const preset = getActivePreset();
  const currentCharacter = getCurrentChatCharacter();

  const activeProvider = getActiveProvider();
  const filteredMessages = messages
    .filter((msg) => msg.role === "user" || msg.role === "assistant")
    .map((msg) => {
      if (activeProvider?.includeThinkingInPrompt && msg.thinkingSummary) {
        return {
          ...msg,
          content: `${msg.content}\n\n[Thinking summary]\n${msg.thinkingSummary}`,
        };
      }
      return msg;
    });

  const loreContext = resolveLorebookPromptContext({ messages: filteredMessages, character: currentCharacter });
  const systemPrompt = buildPresetSystemPrompt({ preset, character: currentCharacter, loreContext });
  const vars = buildPromptVariables(currentCharacter, systemPrompt, loreContext);

  if (shouldApplyInstructPreset(preset)) {
    return buildInstructTranscriptMessages({ messages: filteredMessages, preset, systemPrompt });
  }

  const injected = buildPresetInjectedMessages({ preset, vars });
  const apiMessages = [];

  if (systemPrompt) {
    apiMessages.push({ role: "system", content: systemPrompt });
  }

  injected.before.forEach((msg) => apiMessages.push(msg));

  for (const msg of filteredMessages) {
    apiMessages.push({ role: msg.role, content: msg.content });
  }

  injected.after.forEach((msg) => apiMessages.push(msg));

  return apiMessages;
}

function shouldApplyInstructPreset(preset) {
  return Boolean(
    preset?.instruct?.inputSequence
    || preset?.instruct?.outputSequence
    || preset?.instruct?.systemSequence
    || preset?.source === "sillytavern-instruct"
  );
}

function buildInstructTranscriptMessages({ messages, preset, systemPrompt }) {
  const transcript = buildInstructTranscript({ messages, preset, systemPrompt });
  return [{ role: "user", content: transcript }];
}

function buildInstructTranscript({ messages, preset, systemPrompt }) {
  const instruct = preset?.instruct || {};
  const transcriptParts = [];
  const dialogue = messages.filter((msg) => msg.role === "user" || msg.role === "assistant");

  if (systemPrompt) {
    transcriptParts.push(formatInstructMessageContent({
      role: "system",
      content: systemPrompt,
      index: 0,
      total: dialogue.length + 1,
      instruct,
    }));
  }

  dialogue.forEach((message, idx) => {
    transcriptParts.push(formatInstructMessageContent({
      role: message.role,
      content: message.content,
      index: idx + 1,
      total: dialogue.length + 1,
      instruct,
    }));
  });

  const assistantCue = buildAssistantContinuationCue(instruct);
  if (assistantCue) {
    transcriptParts.push(assistantCue);
  }

  return transcriptParts.filter(Boolean).join(instruct.wrap ? "\n" : "\n\n").trim();
}

function buildAssistantContinuationCue(instruct) {
  const sequenceRaw = instruct.lastOutputSequence || instruct.outputSequence || "";
  const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
  if (!sequence) return "";
  return `${sequence}${sequence ? "\n" : ""}`;
}

function formatInstructMessageContent({ role, content, index, total, instruct }) {
  const isFirstNonSystem = index === 1;
  const isLast = index === total - 1;
  const stopSequence = applyMacrosToText(instruct.stopSequence || "", getCurrentChatCharacter());
  const resolvedContent = applyMacrosToText(String(content || ""), getCurrentChatCharacter());

  if (role === "system") {
    const sequenceRaw = instruct.systemSameAsUser
      ? (instruct.firstInputSequence || instruct.inputSequence || "")
      : (instruct.systemSequence || instruct.storyStringPrefix || "");
    const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
    const suffix = applyMacrosToText(instruct.systemSuffix || instruct.storyStringSuffix || stopSequence || "", getCurrentChatCharacter());
    return `${sequence}${sequence ? "\n" : ""}${resolvedContent}${suffix}`.trimEnd();
  }

  if (role === "user") {
    const sequenceRaw = isFirstNonSystem
      ? (instruct.firstInputSequence || instruct.inputSequence || "")
      : isLast
        ? (instruct.lastInputSequence || instruct.inputSequence || "")
        : (instruct.inputSequence || "");
    const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
    const suffix = applyMacrosToText(instruct.inputSuffix || stopSequence || "", getCurrentChatCharacter());
    return `${sequence}${sequence ? "\n" : ""}${resolvedContent}${suffix}`.trimEnd();
  }

  if (role === "assistant") {
    const sequenceRaw = isFirstNonSystem
      ? (instruct.firstOutputSequence || instruct.outputSequence || "")
      : isLast
        ? (instruct.lastOutputSequence || instruct.outputSequence || "")
        : (instruct.outputSequence || "");
    const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
    const suffix = applyMacrosToText(instruct.outputSuffix || stopSequence || "", getCurrentChatCharacter());
    return `${sequence}${sequence ? "\n" : ""}${resolvedContent}${suffix}`.trimEnd();
  }

  return resolvedContent;
}

function buildPresetSystemPrompt({ preset, character, loreContext = null }) {
  const includeSystemPrompt = preset?.useSystemPrompt !== false;
  const fallbackPrompt = includeSystemPrompt ? (preset?.systemPrompt?.trim() || state.systemPrompt?.trim() || "") : "";
  const vars = buildPromptVariables(character, fallbackPrompt, loreContext);
  const contextBlock = buildContextPresetPrompt({ preset, vars });
  const template = preset?.contextTemplate || defaults.presets[0].contextTemplate;
  const templateHandlesLore = /{{\s*(wiBefore|wiAfter|anchorBefore|anchorAfter)\s*}}/i.test(template);

  const parts = [fallbackPrompt];
  if (!templateHandlesLore && vars.wiBefore) parts.push(vars.wiBefore);
  if (contextBlock) parts.push(contextBlock);
  if (!templateHandlesLore && vars.wiAfter) parts.push(vars.wiAfter);

  return parts.filter(Boolean).join("\n\n").trim();
}

function buildPresetInjectedMessages({ preset, vars }) {
  const before = [];
  const after = [];

  getOrderedPromptEntries(preset)
    .filter((prompt) => prompt._orderEnabled !== false && prompt.enabled !== false)
    .forEach((prompt) => {
      const rendered = renderOpenAiPromptEntry(prompt, vars);
      if (!rendered) return;
      const role = ["system", "user", "assistant"].includes(prompt.role) ? prompt.role : "system";
      const target = Number(prompt.injection_position ?? 0) === 1 ? after : before;
      target.push({ role, content: rendered });
    });

  return { before, after };
}

function buildPromptVariables(character, systemPrompt = "", loreContext = null) {
  const name = character?.name?.trim() || "Character";
  const userName = resolveCurrentUserName();
  const preset = getActivePreset();
  const scenario = character?.scenario?.trim() || character?.data?.scenario?.trim() || "";
  const personality = character?.personality?.trim() || character?.data?.personality?.trim() || character?.data?.description?.trim() || "";
  const greeting = character?.greeting?.trim() || character?.data?.first_mes?.trim() || "";
  const description = character?.data?.description?.trim() || personality || "";
  const creatorNotes = character?.data?.creator_notes || character?.data?.creatorNotes || "";
  const charVersion = character?.data?.character_version || character?.data?.char_version || character?.data?.version || "";
  const mesExamplesRaw = character?.data?.mes_example || character?.data?.mesExamples || "";
  const loreBeforeRaw = String(loreContext?.beforeText || "").trim();
  const loreAfterRaw = String(loreContext?.afterText || "").trim();

  const vars = {
    char: name,
    user: userName,
    system: systemPrompt,
    systemPrompt,
    defaultSystemPrompt: systemPrompt,
    instructSystem: systemPrompt,
    instructSystemPrompt: systemPrompt,
    description,
    personality,
    scenario,
    persona: "",
    anchorBefore: loreBeforeRaw,
    anchorAfter: loreAfterRaw,
    wiBefore: loreBeforeRaw,
    wiAfter: loreAfterRaw,
    greeting,
    charDescription: description,
    charPersonality: personality,
    charScenario: scenario,
    charFirstMessage: greeting,
    charCreatorNotes: creatorNotes,
    creatorNotes,
    charVersion,
    version: charVersion,
    char_version: charVersion,
    mesExamplesRaw,
    mesExamples: mesExamplesRaw,
    chatStart: preset?.chatStart || "",
    exampleSeparator: preset?.exampleSeparator || "",
    chatSeparator: preset?.exampleSeparator || "",
  };

  const resolvedGreeting = applyMacrosToText(greeting, character, userName);
  const resolvedWiBefore = applyMacrosToText(loreBeforeRaw, character, userName, { vars });
  const resolvedWiAfter = applyMacrosToText(loreAfterRaw, character, userName, { vars });

  return {
    ...vars,
    greeting: resolvedGreeting,
    wiBefore: resolvedWiBefore,
    wiAfter: resolvedWiAfter,
    anchorBefore: resolvedWiBefore,
    anchorAfter: resolvedWiAfter,
  };
}

function getOrderedPromptEntries(preset) {
  const promptOrder = Array.isArray(preset?.promptOrder) ? preset.promptOrder : [];
  const prompts = normalizePresetPromptList(Array.isArray(preset?.prompts) ? preset.prompts : []);

  const orderBlock = promptOrder.find((item) => Array.isArray(item?.order));
  if (!orderBlock) {
    return prompts.map((prompt) => ({ ...prompt, _orderEnabled: prompt.enabled !== false }));
  }

  const map = new Map(prompts.map((prompt) => [prompt.identifier, prompt]));
  const ordered = [];

  orderBlock.order.forEach((entry) => {
    const prompt = map.get(entry?.identifier);
    if (!prompt) return;
    ordered.push({ ...prompt, _orderEnabled: entry?.enabled !== false });
  });

  prompts.forEach((prompt) => {
    if (ordered.some((item) => item.identifier === prompt.identifier)) return;
    ordered.push({ ...prompt, _orderEnabled: prompt.enabled !== false });
  });

  return ordered;
}

function buildOpenAiPresetPrompt({ preset, vars, positionFilter = null }) {
  const ordered = getOrderedPromptEntries(preset)
    .filter((prompt) => prompt._orderEnabled !== false && prompt.enabled !== false)
    .filter((prompt) => {
      if (positionFilter === null) return true;
      return Number(prompt.injection_position ?? 0) === Number(positionFilter);
    });

  const rendered = ordered
    .map((prompt) => renderOpenAiPromptEntry(prompt, vars))
    .filter(Boolean);

  return rendered.join("\n\n").trim();
}

function renderOpenAiPromptEntry(prompt, vars) {
  const identifier = prompt.identifier || "";
  if (prompt.marker) {
    if (identifier === "charDescription") return renderTemplate(vars.description, vars);
    if (identifier === "charPersonality") return renderTemplate(vars.personality, vars);
    if (identifier === "scenario") return renderTemplate(vars.scenario, vars);
    if (identifier === "personaDescription") return renderTemplate(vars.persona, vars);
    if (identifier === "main") return renderTemplate(vars.system, vars);
    return "";
  }
  return renderTemplate(prompt.content || "", vars);
}

function buildContextPresetPrompt({ preset, vars }) {
  const template = preset?.contextTemplate || defaults.presets[0].contextTemplate;
  const formattedScenario = preset?.scenarioFormat ? renderTemplate(preset.scenarioFormat, vars) : vars.scenario;
  const formattedPersonality = preset?.personalityFormat ? renderTemplate(preset.personalityFormat, vars) : vars.personality;
  return renderTemplate(template, {
    ...vars,
    scenario: formattedScenario,
    personality: formattedPersonality,
  }).trim();
}

function renderTemplate(template, vars) {
  let output = String(template || "");
  output = output.replace(/{{#if\s+([a-zA-Z0-9_]+)}}([\s\S]*?){{\/if}}/g, (_, key, content) => vars?.[key] ? content : "");
  Object.entries(vars || {}).forEach(([key, value]) => {
    output = output.replace(new RegExp(`{{\\s*${escapeRegex(key)}\\s*}}`, "g"), String(value || ""));
  });
  output = applyMacrosToText(output, null, vars?.user || "", { vars });
  output = output.replace(/\n{3,}/g, "\n\n");
  return output.trim();
}

function estimateTokenCount(text = "") {
  const normalized = String(text || "").trim();
  if (!normalized) return 0;
  return Math.ceil(normalized.length / 4);
}

function normalizeStringArray(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }
  if (typeof value === "string") {
    return value
      .split(/[,\n]/)
      .map((item) => item.trim())
      .filter(Boolean);
  }
  return [];
}

function normalizeLorebookEntry(entry = {}, index = 0) {
  const keys = normalizeStringArray(
    entry.keys
    || entry.key
    || entry.keywords
    || entry.primary_keys
    || entry.triggers
    || entry.match
    || []
  );
  const secondaryKeys = normalizeStringArray(
    entry.secondary_keys
    || entry.secondaryKeys
    || entry.secondary
    || entry.secondary_keywords
    || entry.secondaryKeywords
    || entry.keysecondary
    || entry.key_secondary
    || []
  );
  const content = String(
    entry.content
    ?? entry.text
    ?? entry.value
    ?? entry.entry
    ?? entry.prompt
    ?? entry.desc
    ?? entry.description
    ?? ""
  ).trim();
  const insertionOrder = Number.isFinite(Number(entry.insertion_order ?? entry.insertionOrder ?? entry.order ?? entry.uid))
    ? Number(entry.insertion_order ?? entry.insertionOrder ?? entry.order ?? entry.uid)
    : index;
  const priority = Number.isFinite(Number(entry.priority ?? entry.weight)) ? Number(entry.priority ?? entry.weight) : 100;
  const probabilityRaw = Number(entry.probability ?? entry.chance ?? 100);
  const probability = Math.min(100, Math.max(0, Number.isFinite(probabilityRaw) ? probabilityRaw : 100));

  let selectiveLogic = String(entry.selectiveLogic || entry.selective_logic || entry.logic || "AND").trim().toUpperCase();
  if (!["AND", "OR", "NOT"].includes(selectiveLogic)) selectiveLogic = "AND";

  return {
    id: entry.id || entry.uid || crypto.randomUUID(),
    title: String(entry.title || entry.name || entry.comment || `Entry ${index + 1}`),
    keys,
    secondaryKeys,
    content,
    insertionOrder,
    priority,
    caseSensitive: Boolean(entry.caseSensitive ?? entry.case_sensitive ?? false),
    selective: Boolean(entry.selective ?? secondaryKeys.length > 0),
    selectiveLogic,
    constant: Boolean(entry.constant ?? false),
    probability,
    enabled: entry.enabled !== false,
  };
}

function parseLorebookEntries(raw = {}) {
  if (!raw || typeof raw !== "object") return [];

  const candidateSources = [
    raw.entries,
    raw.data?.entries,
    raw.world_info,
    raw.world_info?.entries,
    raw.worldInfo,
    raw.worldInfo?.entries,
    raw.lorebook?.entries,
    raw.book?.entries,
    raw.items,
  ];

  const seen = new Set();
  const entriesSource = [];

  candidateSources.forEach((source) => {
    if (!source) return;

    if (Array.isArray(source)) {
      source.forEach((item) => {
        if (!item || typeof item !== "object") return;
        if (seen.has(item)) return;
        seen.add(item);
        entriesSource.push(item);
      });
      return;
    }

    if (typeof source === "object") {
      Object.values(source).forEach((item) => {
        if (!item || typeof item !== "object") return;
        if (seen.has(item)) return;
        seen.add(item);
        entriesSource.push(item);
      });
    }
  });

  return entriesSource
    .map((entry, index) => normalizeLorebookEntry(entry, index))
    .filter((entry) => {
      const hasKeys = Array.isArray(entry.keys) && entry.keys.length > 0;
      const hasSecondary = Array.isArray(entry.secondaryKeys) && entry.secondaryKeys.length > 0;
      const hasTitle = String(entry.title || "").trim().length > 0;
      const hasContent = String(entry.content || "").trim().length > 0;
      return hasContent || hasKeys || hasSecondary || hasTitle || entry.constant;
    });
}

function normalizeLorebookRecord(item = {}) {
  const fromData = item?.data && typeof item.data === "object" ? item.data : {};
  const source = {
    ...(fromData && typeof fromData === "object" ? fromData : {}),
    ...(item && typeof item === "object" ? item : {}),
  };

  if (Object.prototype.hasOwnProperty.call(item || {}, "entries")) {
    source.entries = item.entries;
  }
  if (Object.prototype.hasOwnProperty.call(item || {}, "world_info")) {
    source.world_info = item.world_info;
  }

  const entries = parseLorebookEntries(source);
  const scanDepthRaw = Number(source.scanDepth ?? source.scan_depth ?? source.depth ?? defaults.lorebookSettings.scanDepth);
  const tokenBudgetRaw = Number(source.tokenBudget ?? source.token_budget ?? source.budget ?? defaults.lorebookSettings.tokenBudget);

  return {
    id: item.id || source.id || crypto.randomUUID(),
    name: String(source.name || source.title || item.name || "Imported lorebook"),
    description: String(source.description || item.description || ""),
    scanDepth: Math.min(64, Math.max(1, Number.isFinite(scanDepthRaw) ? scanDepthRaw : defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.isFinite(tokenBudgetRaw) ? tokenBudgetRaw : defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: Boolean(source.recursiveScanning ?? source.recursive_scanning ?? source.recursive ?? defaults.lorebookSettings.recursiveScanning),
    enabled: item.enabled !== false,
    entries,
    data: {
      ...(source && typeof source === "object" ? source : {}),
      entries,
    },
  };
}

function matchesKeyword(text, keyword, caseSensitive = false) {
  const source = String(text || "");
  const needle = String(keyword || "").trim();
  if (!needle) return false;
  const flags = caseSensitive ? "g" : "gi";
  const pattern = `(^|[^A-Za-z0-9_])${escapeRegex(needle)}(?=$|[^A-Za-z0-9_])`;
  return new RegExp(pattern, flags).test(source);
}

function resolveLorebookPromptContext({ messages = [], character = null } = {}) {
  const chat = getCurrentChat();
  const attachedIds = new Set(Array.isArray(chat?.lorebookIds) ? chat.lorebookIds : []);
  const enabledLorebooks = state.lorebooks.filter((book) => attachedIds.has(book.id));
  if (!enabledLorebooks.length) return { beforeText: "", afterText: "", entries: [] };

  const activeSettings = getActiveChatLorebookSettings();
  const defaultScanDepth = activeSettings.scanDepth;
  const defaultTokenBudget = activeSettings.tokenBudget;
  const recursiveEnabled = activeSettings.recursiveScanning;

  const history = Array.isArray(messages)
    ? messages.filter((msg) => msg && (msg.role === "user" || msg.role === "assistant")).slice(-defaultScanDepth)
    : [];
  const characterGreeting = String(character?.greeting || character?.data?.first_mes || "").trim();
  const characterSummary = [
    String(character?.name || "").trim(),
    String(character?.scenario || character?.data?.scenario || "").trim(),
    String(character?.personality || character?.data?.personality || character?.data?.description || "").trim(),
    characterGreeting,
  ].filter(Boolean).join("\n");
  let searchableText = [
    characterSummary,
    history.map((msg) => String(msg.content || "")).join("\n"),
  ].filter(Boolean).join("\n");

  const selectedEntries = [];
  const selectedIds = new Set();

  const shouldActivateEntry = (entry, searchText) => {
    if (!entry || entry.enabled === false) return false;
    if (entry.constant) return true;

    const primaryMatches = entry.keys.some((key) => matchesKeyword(searchText, key, entry.caseSensitive));
    if (!primaryMatches) return false;

    if (!entry.selective || !entry.secondaryKeys.length) {
      return true;
    }

    const secondaryMatches = entry.secondaryKeys.some((key) => matchesKeyword(searchText, key, entry.caseSensitive));

    if (entry.selectiveLogic === "NOT") return !secondaryMatches;
    if (entry.selectiveLogic === "OR") return primaryMatches || secondaryMatches;
    return primaryMatches && secondaryMatches;
  };

  const runPass = () => {
    let added = false;
    for (const lorebook of enabledLorebooks) {
      const entries = Array.isArray(lorebook.entries) ? lorebook.entries : [];
      for (const entry of entries) {
        if (selectedIds.has(entry.id)) continue;
        if (!shouldActivateEntry(entry, searchableText)) continue;

        const chanceRoll = Math.random() * 100;
        if (entry.probability < 100 && chanceRoll > entry.probability) continue;

        selectedIds.add(entry.id);
        selectedEntries.push({
          ...entry,
          _lorebookId: lorebook.id,
          _lorebookName: lorebook.name,
          _effectiveTokenBudget: Math.min(defaultTokenBudget, lorebook.tokenBudget || defaultTokenBudget),
        });
        searchableText += `\n${entry.content}`;
        added = true;
      }
    }
    return added;
  };

  runPass();
  if (recursiveEnabled) {
    for (let pass = 0; pass < 8; pass += 1) {
      if (!runPass()) break;
    }
  }

  const sorted = selectedEntries
    .slice()
    .sort((a, b) => {
      const byPriority = Number(b.priority || 0) - Number(a.priority || 0);
      if (byPriority !== 0) return byPriority;
      const byOrder = Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0);
      if (byOrder !== 0) return byOrder;
      return String(a.title || "").localeCompare(String(b.title || ""));
    });

  const kept = [];
  let usedTokens = 0;
  const globalBudget = defaultTokenBudget;
  for (const entry of sorted) {
    const text = applyMacrosToText(String(entry.content || "").trim(), character);
    if (!text) continue;
    const estimated = estimateTokenCount(text);
    if (usedTokens + estimated > globalBudget) continue;
    kept.push({ ...entry, _rendered: text, _estimatedTokens: estimated });
    usedTokens += estimated;
  }

  kept.sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0));
  const beforeText = kept.map((entry) => entry._rendered).join("\n\n").trim();

  return {
    beforeText,
    afterText: "",
    entries: kept,
    estimatedTokens: usedTokens,
    scanDepthUsed: defaultScanDepth,
    tokenBudgetUsed: defaultTokenBudget,
  };
}

function syncLorebookSettingsInputs() {
  if (els.lorebookScanDepthInput) {
    els.lorebookScanDepthInput.value = String(state.lorebookSettings.scanDepth || defaults.lorebookSettings.scanDepth);
  }
  if (els.lorebookTokenBudgetInput) {
    els.lorebookTokenBudgetInput.value = String(state.lorebookSettings.tokenBudget || defaults.lorebookSettings.tokenBudget);
  }
  if (els.lorebookRecursiveInput) {
    els.lorebookRecursiveInput.checked = state.lorebookSettings.recursiveScanning !== false;
  }
  setInlineStatus(els.lorebookSettingsStatus, "");
}

function getLabeledLorebookSettings(book = null) {
  const fallback = {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(state.lorebookSettings?.scanDepth, 10) || defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(state.lorebookSettings?.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: state.lorebookSettings?.recursiveScanning !== false,
  };

  if (!book || typeof book !== "object") return fallback;

  return {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(book.scanDepth, 10) || fallback.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(book.tokenBudget, 10) || fallback.tokenBudget)),
    recursiveScanning: book.recursiveScanning ?? fallback.recursiveScanning,
  };
}

function saveLorebookSettingsFromInputs() {
  const scanDepth = Math.min(64, Math.max(1, Number.parseInt(els.lorebookScanDepthInput?.value || "", 10) || defaults.lorebookSettings.scanDepth));
  const tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(els.lorebookTokenBudgetInput?.value || "", 10) || defaults.lorebookSettings.tokenBudget));

  state.lorebookSettings.scanDepth = scanDepth;
  state.lorebookSettings.tokenBudget = tokenBudget;
  state.lorebookSettings.recursiveScanning = Boolean(els.lorebookRecursiveInput?.checked ?? true);

  persistState();
  syncLorebookSettingsInputs();
  setInlineStatus(els.lorebookSettingsStatus, "Lore settings saved");
  showToast("Lore settings saved", "success");
}

function setActiveLibraryPanel(panelId = "libraryLorebooksPanel") {
  const target = panelId || "libraryLorebooksPanel";
  els.libraryPanels?.forEach((panel) => {
    panel.classList.toggle("hidden", panel.id !== target);
    panel.classList.toggle("active", panel.id === target);
  });
  els.libraryCategoryChips?.forEach((chip) => {
    chip.classList.toggle("active", chip.dataset.libraryPanel === target);
  });
}

function parseCommaSeparatedList(value = "") {
  return String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function openLorebookEntryEditor(index = null) {
  if (!els.lorebookEntryEditorCard) return;

  const hasIndex = Number.isInteger(index) && index >= 0 && lorebookEntryDraft[index];
  const entry = hasIndex ? lorebookEntryDraft[index] : {
    title: "",
    keys: [],
    secondaryKeys: [],
    content: "",
    priority: 100,
    probability: 100,
    caseSensitive: false,
    selective: false,
    selectiveLogic: "AND",
    constant: false,
    enabled: true,
  };

  els.lorebookEntryEditorCard.classList.remove("hidden");
  if (hasIndex) {
    els.lorebookEntryEditorCard.dataset.editingIndex = String(index);
  } else {
    delete els.lorebookEntryEditorCard.dataset.editingIndex;
  }

  if (els.lorebookEntryTitleInput) els.lorebookEntryTitleInput.value = entry.title || "";
  if (els.lorebookEntryKeysInput) els.lorebookEntryKeysInput.value = Array.isArray(entry.keys) ? entry.keys.join(", ") : "";
  if (els.lorebookEntrySecondaryKeysInput) els.lorebookEntrySecondaryKeysInput.value = Array.isArray(entry.secondaryKeys) ? entry.secondaryKeys.join(", ") : "";
  if (els.lorebookEntryContentInput) els.lorebookEntryContentInput.value = entry.content || entry.entry || "";
  if (els.lorebookEntryPriorityInput) els.lorebookEntryPriorityInput.value = String(Number.isFinite(Number(entry.priority)) ? Number(entry.priority) : 100);
  if (els.lorebookEntryProbabilityInput) els.lorebookEntryProbabilityInput.value = String(Number.isFinite(Number(entry.probability)) ? Number(entry.probability) : 100);
  if (els.lorebookEntryEnabledInput) els.lorebookEntryEnabledInput.checked = entry.enabled !== false;
  if (els.lorebookEntryConstantInput) els.lorebookEntryConstantInput.checked = Boolean(entry.constant);
  if (els.lorebookEntryCaseSensitiveInput) els.lorebookEntryCaseSensitiveInput.checked = Boolean(entry.caseSensitive);
  if (els.lorebookEntrySelectiveInput) els.lorebookEntrySelectiveInput.checked = Boolean(entry.selective);
  if (els.lorebookEntrySelectiveLogicInput) els.lorebookEntrySelectiveLogicInput.value = entry.selectiveLogic || "AND";

  if (els.lorebookEntryDeleteBtn) {
    els.lorebookEntryDeleteBtn.classList.toggle("hidden", !hasIndex);
  }

  els.lorebookEntryTitleInput?.focus();
}

function closeLorebookEntryEditor() {
  if (!els.lorebookEntryEditorCard) return;
  els.lorebookEntryEditorCard.classList.add("hidden");
  delete els.lorebookEntryEditorCard.dataset.editingIndex;
  if (els.lorebookEntryTitleInput) els.lorebookEntryTitleInput.value = "";
  if (els.lorebookEntryKeysInput) els.lorebookEntryKeysInput.value = "";
  if (els.lorebookEntrySecondaryKeysInput) els.lorebookEntrySecondaryKeysInput.value = "";
  if (els.lorebookEntryContentInput) els.lorebookEntryContentInput.value = "";
  if (els.lorebookEntryPriorityInput) els.lorebookEntryPriorityInput.value = "100";
  if (els.lorebookEntryProbabilityInput) els.lorebookEntryProbabilityInput.value = "100";
  if (els.lorebookEntryEnabledInput) els.lorebookEntryEnabledInput.checked = true;
  if (els.lorebookEntryConstantInput) els.lorebookEntryConstantInput.checked = false;
  if (els.lorebookEntryCaseSensitiveInput) els.lorebookEntryCaseSensitiveInput.checked = false;
  if (els.lorebookEntrySelectiveInput) els.lorebookEntrySelectiveInput.checked = false;
  if (els.lorebookEntrySelectiveLogicInput) els.lorebookEntrySelectiveLogicInput.value = "AND";
}

function saveLorebookEntryFromEditor(options = {}) {
  const skipClose = Boolean(options?.skipClose);
  const editingIndex = Number.parseInt(els.lorebookEntryEditorCard?.dataset?.editingIndex || "-1", 10);
  const fallbackTitle = Number.isFinite(editingIndex) && editingIndex >= 0
    ? `Entry ${editingIndex + 1}`
    : `Entry ${lorebookEntryDraft.length + 1}`;
  const title = els.lorebookEntryTitleInput?.value?.trim() || fallbackTitle;
  const keys = parseCommaSeparatedList(els.lorebookEntryKeysInput?.value || "");
  const secondaryKeys = parseCommaSeparatedList(els.lorebookEntrySecondaryKeysInput?.value || "");
  const content = els.lorebookEntryContentInput?.value?.trim() || "";

  const payload = normalizeLorebookEntry({
    id: crypto.randomUUID(),
    title,
    keys,
    secondary_keys: secondaryKeys,
    content,
    priority: Number.parseInt(els.lorebookEntryPriorityInput?.value || "100", 10),
    probability: Number.parseInt(els.lorebookEntryProbabilityInput?.value || "100", 10),
    case_sensitive: Boolean(els.lorebookEntryCaseSensitiveInput?.checked),
    selective: Boolean(els.lorebookEntrySelectiveInput?.checked),
    selective_logic: els.lorebookEntrySelectiveLogicInput?.value || "AND",
    constant: Boolean(els.lorebookEntryConstantInput?.checked),
    enabled: Boolean(els.lorebookEntryEnabledInput?.checked ?? true),
  }, 0);

  if (Number.isFinite(editingIndex) && editingIndex >= 0 && lorebookEntryDraft[editingIndex]) {
    lorebookEntryDraft[editingIndex] = {
      ...lorebookEntryDraft[editingIndex],
      ...payload,
      id: lorebookEntryDraft[editingIndex].id || payload.id,
    };
    setInlineStatus(els.lorebookEditorStatus, `Updated entry ${editingIndex + 1}`);
  } else {
    lorebookEntryDraft.push(payload);
    setInlineStatus(els.lorebookEditorStatus, "Entry added");
  }

  lorebookEntryDraft = lorebookEntryDraft.map((entry, index) => ({ ...entry, insertionOrder: index }));
  lorebookEntrySearchQuery = "";
  if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
  if (!skipClose) {
    closeLorebookEntryEditor();
  }
  renderLorebookEntryDraft();
  return true;
}

function deleteLorebookEntryFromEditor() {
  const editingIndex = Number.parseInt(els.lorebookEntryEditorCard?.dataset?.editingIndex || "-1", 10);
  if (!Number.isFinite(editingIndex) || editingIndex < 0 || !lorebookEntryDraft[editingIndex]) {
    closeLorebookEntryEditor();
    return;
  }

  lorebookEntryDraft.splice(editingIndex, 1);
  lorebookEntryDraft = lorebookEntryDraft.map((entry, index) => ({ ...entry, insertionOrder: index }));
  closeLorebookEntryEditor();
  renderLorebookEntryDraft();
  setInlineStatus(els.lorebookEditorStatus, "Entry deleted");
}

function showLorebookForm(lorebook = null) {
  const sourceLorebook = lorebook?.id ? (state.lorebooks.find((book) => book.id === lorebook.id) || lorebook) : lorebook;
  const normalized = sourceLorebook ? normalizeLorebookRecord(sourceLorebook) : normalizeLorebookRecord({ name: "", description: "", entries: [] });
  if (els.lorebookEditorView) {
    els.lorebookEditorView.dataset.editingId = normalized.id || lorebook?.id || "";
  }

  if (els.lorebookNameInput) els.lorebookNameInput.value = normalized.name || "";
  if (els.lorebookDescriptionInput) els.lorebookDescriptionInput.value = normalized.description || "";
  lorebookEntryDraft = (Array.isArray(normalized.entries) ? normalized.entries : [])
    .slice()
    .sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0))
    .map((entry, index) => ({ ...entry, insertionOrder: index }));

  lorebookEntrySearchQuery = "";
  if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
  closeLorebookEntryEditor();
  renderLorebookEntryDraft();
  setInlineStatus(els.lorebookEditorStatus, "");
  switchView("lorebookEditorView");
  els.lorebookNameInput?.focus();
}

function hideLorebookForm() {
  if (els.lorebookEditorView) {
    delete els.lorebookEditorView.dataset.editingId;
  }
  if (els.lorebookNameInput) els.lorebookNameInput.value = "";
  if (els.lorebookDescriptionInput) els.lorebookDescriptionInput.value = "";
  lorebookEntryDraft = [];
  lorebookEntrySearchQuery = "";
  if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
  closeLorebookEntryEditor();
  renderLorebookEntryDraft();
  setInlineStatus(els.lorebookEditorStatus, "");
  switchView("libraryView");
  setActiveLibraryPanel("libraryLorebooksPanel");
}

function clearLorebookEntryLongPressState() {
  if (lorebookEntryLongPress.timer) {
    clearTimeout(lorebookEntryLongPress.timer);
  }
  if (lorebookEntryLongPress.moveHandler) {
    document.removeEventListener("pointermove", lorebookEntryLongPress.moveHandler);
  }
  if (lorebookEntryLongPress.upHandler) {
    document.removeEventListener("pointerup", lorebookEntryLongPress.upHandler);
    document.removeEventListener("pointercancel", lorebookEntryLongPress.upHandler);
  }
  if (lorebookEntryLongPress.autoScrollRaf) {
    cancelAnimationFrame(lorebookEntryLongPress.autoScrollRaf);
  }

  els.lorebookEntryList?.classList.remove("drag-mode");
  els.lorebookEntryList?.querySelectorAll(".lorebook-entry-item").forEach((node) => {
    node.classList.remove("drag-source", "drop-target", "longpress-arming", "longpress-active");
    node.style.transform = "";
  });

  lorebookEntryLongPress = {
    timer: null,
    pointerId: null,
    sourceIndex: -1,
    targetIndex: -1,
    active: false,
    moveHandler: null,
    upHandler: null,
    autoScrollRaf: null,
    latestClientX: 0,
    latestClientY: 0,
    scrollContainer: null,
  };
}

function isScrollableY(element) {
  if (!element || !(element instanceof Element)) return false;
  const style = getComputedStyle(element);
  const allowsScroll = /(auto|scroll|overlay)/.test(style.overflowY || "");
  return allowsScroll && element.scrollHeight > element.clientHeight + 2;
}

function getDragScrollContainer(startElement) {
  let node = startElement;
  while (node && node !== document.body) {
    if (isScrollableY(node)) return node;
    node = node.parentElement;
  }
  return document.scrollingElement || document.documentElement;
}

function getScrollContainerRect(container) {
  if (!container) {
    return { top: 0, bottom: window.innerHeight };
  }
  if (container === document.scrollingElement || container === document.documentElement || container === document.body) {
    return { top: 0, bottom: window.innerHeight };
  }
  return container.getBoundingClientRect();
}

function applyDragAutoScroll(container, clientY) {
  if (!container) return;

  const rect = getScrollContainerRect(container);
  const edgeSize = 72;
  const maxStep = 18;
  let delta = 0;

  if (clientY < rect.top + edgeSize) {
    const t = Math.max(0, (rect.top + edgeSize - clientY) / edgeSize);
    delta = -Math.ceil(maxStep * t);
  } else if (clientY > rect.bottom - edgeSize) {
    const t = Math.max(0, (clientY - (rect.bottom - edgeSize)) / edgeSize);
    delta = Math.ceil(maxStep * t);
  }

  if (!delta) return;

  if (container === document.scrollingElement || container === document.documentElement || container === document.body) {
    window.scrollBy(0, delta);
    return;
  }

  container.scrollTop += delta;
}

function renderLorebookEntryDraft() {
  if (!els.lorebookEntryList) return;

  clearLorebookEntryLongPressState();
  els.lorebookEntryList.innerHTML = "";

  const query = lorebookEntrySearchQuery.trim().toLowerCase();
  const visibleEntries = lorebookEntryDraft
    .map((entry, index) => ({ entry, index }))
    .filter(({ entry }) => {
      if (!query) return true;
      const haystack = `${entry.title || ""} ${(entry.keys || []).join(" ")} ${(entry.secondaryKeys || []).join(" ")} ${entry.content || ""}`.toLowerCase();
      return haystack.includes(query);
    });

  if (!visibleEntries.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">${query ? "No entries match your search." : "No entries yet. Create one."}</div>`;
    els.lorebookEntryList.appendChild(empty);
    return;
  }

  visibleEntries.forEach(({ entry, index }) => {
    const card = document.createElement("article");
    card.className = "entity-card lorebook-entry-item";
    card.dataset.entryIndex = String(index);

    const keyPreview = Array.isArray(entry.keys) ? entry.keys.slice(0, 4).join(", ") : "";
    const contentText = String(entry.content || entry.entry || "").trim();
    const contentPreview = contentText.slice(0, 260);

    card.innerHTML = `
      <div class="entity-card-header">
        <h3>Entry ${index + 1}: ${escapeHtml(entry.title || "Untitled")}</h3>
        <span class="entity-meta">${(Array.isArray(entry.keys) ? entry.keys.length : 0)} keys</span>
      </div>
      <div class="entity-meta">${escapeHtml(keyPreview || "No keys")}</div>
      <div class="entity-meta">${escapeHtml(contentPreview || "No content")}</div>
      <div class="entity-card-actions">
        <button class="secondary" type="button" data-action="edit-entry">Edit</button>
      </div>
    `;

    card.querySelector('[data-action="edit-entry"]')?.addEventListener("click", (event) => {
      event.stopPropagation();
      openLorebookEntryEditor(index);
    });

    card.addEventListener("pointerdown", (event) => {
      if (event.target?.closest?.("button")) return;
      if (event.button !== undefined && event.button !== 0) return;

      clearLorebookEntryLongPressState();
      card.classList.add("longpress-arming");

      lorebookEntryLongPress.pointerId = event.pointerId;
      lorebookEntryLongPress.sourceIndex = index;
      lorebookEntryLongPress.targetIndex = index;
      lorebookEntryLongPress.latestClientX = event.clientX;
      lorebookEntryLongPress.latestClientY = event.clientY;
      lorebookEntryLongPress.scrollContainer = getDragScrollContainer(els.lorebookEntryList || card);

      const originX = event.clientX;
      const originY = event.clientY;

      const removePreLongPressListeners = () => {
        document.removeEventListener("pointermove", preLongPressMove);
        document.removeEventListener("pointerup", preLongPressEnd);
        document.removeEventListener("pointercancel", preLongPressEnd);
      };

      const preLongPressMove = (moveEvent) => {
        if (moveEvent.pointerId !== lorebookEntryLongPress.pointerId) return;
        lorebookEntryLongPress.latestClientX = moveEvent.clientX;
        lorebookEntryLongPress.latestClientY = moveEvent.clientY;

        // User is scrolling, cancel long-press arming and let normal scroll continue.
        const dx = Math.abs(moveEvent.clientX - originX);
        const dy = Math.abs(moveEvent.clientY - originY);
        if (!lorebookEntryLongPress.active && (dy > 10 || dx > 10)) {
          if (lorebookEntryLongPress.timer) clearTimeout(lorebookEntryLongPress.timer);
          lorebookEntryLongPress.timer = null;
          card.classList.remove("longpress-arming");
          removePreLongPressListeners();
        }
      };

      const preLongPressEnd = (upEvent) => {
        if (upEvent.pointerId !== lorebookEntryLongPress.pointerId) return;
        removePreLongPressListeners();
        if (!lorebookEntryLongPress.active) {
          clearLorebookEntryLongPressState();
        }
      };

      lorebookEntryLongPress.timer = setTimeout(() => {
        lorebookEntryLongPress.active = true;
        card.classList.remove("longpress-arming");
        card.classList.add("drag-source", "longpress-active");
        els.lorebookEntryList?.classList.add("drag-mode");

        lorebookEntryLongPress.moveHandler = (moveEvent) => {
          if (moveEvent.pointerId !== lorebookEntryLongPress.pointerId) return;

          lorebookEntryLongPress.latestClientX = moveEvent.clientX;
          lorebookEntryLongPress.latestClientY = moveEvent.clientY;

          // prevent page pan only after drag mode starts
          moveEvent.preventDefault();

          const dy = moveEvent.clientY - originY;
          card.style.transform = `translateY(${dy}px) scale(1.02)`;

          applyDragAutoScroll(lorebookEntryLongPress.scrollContainer, lorebookEntryLongPress.latestClientY);

          const hovered = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY)?.closest?.(".lorebook-entry-item");
          if (!hovered || hovered.parentElement !== els.lorebookEntryList) return;

          const hoverIndex = Number.parseInt(hovered.dataset.entryIndex || "-1", 10);
          if (!Number.isFinite(hoverIndex) || hoverIndex < 0) return;

          lorebookEntryLongPress.targetIndex = hoverIndex;
          els.lorebookEntryList?.querySelectorAll(".lorebook-entry-item").forEach((itemNode) => itemNode.classList.remove("drop-target"));
          if (hoverIndex !== lorebookEntryLongPress.sourceIndex) {
            hovered.classList.add("drop-target");
          }
        };

        lorebookEntryLongPress.upHandler = (upEvent) => {
          if (upEvent.pointerId !== lorebookEntryLongPress.pointerId) return;

          const from = lorebookEntryLongPress.sourceIndex;
          const to = lorebookEntryLongPress.targetIndex;

          if (
            lorebookEntryLongPress.active
            && Number.isFinite(from)
            && Number.isFinite(to)
            && from >= 0
            && to >= 0
            && from !== to
            && lorebookEntryDraft[from]
          ) {
            const [movedEntry] = lorebookEntryDraft.splice(from, 1);
            const insertionIndex = from < to ? Math.max(0, to - 1) : to;
            lorebookEntryDraft.splice(insertionIndex, 0, movedEntry);
            lorebookEntryDraft = lorebookEntryDraft.map((item, itemIndex) => ({
              ...item,
              insertionOrder: itemIndex,
            }));
            renderLorebookEntryDraft();
            setInlineStatus(els.lorebookEditorStatus, `Moved entry ${from + 1} to ${insertionIndex + 1}`);
          } else {
            clearLorebookEntryLongPressState();
            renderLorebookEntryDraft();
          }
        };

        const autoScrollLoop = () => {
          if (!lorebookEntryLongPress.active) return;
          applyDragAutoScroll(lorebookEntryLongPress.scrollContainer, lorebookEntryLongPress.latestClientY);
          lorebookEntryLongPress.autoScrollRaf = requestAnimationFrame(autoScrollLoop);
        };

        document.addEventListener("pointermove", lorebookEntryLongPress.moveHandler, { passive: false });
        document.addEventListener("pointerup", lorebookEntryLongPress.upHandler);
        document.addEventListener("pointercancel", lorebookEntryLongPress.upHandler);
        lorebookEntryLongPress.autoScrollRaf = requestAnimationFrame(autoScrollLoop);
      }, 320);

      document.addEventListener("pointermove", preLongPressMove);
      document.addEventListener("pointerup", preLongPressEnd);
      document.addEventListener("pointercancel", preLongPressEnd);
    });

    card.addEventListener("contextmenu", (event) => event.preventDefault());
    els.lorebookEntryList.appendChild(card);
  });
}

function tryParseLooseJson(text = "") {
  const source = String(text || "").trim();
  if (!source) throw new Error("Empty file");

  try {
    return JSON.parse(source);
  } catch {
    // continue
  }

  // Remove dangling commas first.
  try {
    const noTrailingCommas = source.replace(/,\s*([}\]])/g, "$1");
    return JSON.parse(noTrailingCommas);
  } catch {
    // continue
  }

  // If the file is truncated, try to salvage by balancing braces/brackets.
  try {
    let candidate = source
      .replace(/,\s*([}\]])/g, "$1")
      .replace(/\u0000/g, "")
      .trim();

    let inString = false;
    let escape = false;
    let curly = 0;
    let square = 0;

    for (let i = 0; i < candidate.length; i += 1) {
      const ch = candidate[i];
      if (escape) {
        escape = false;
        continue;
      }
      if (ch === "\\") {
        escape = true;
        continue;
      }
      if (ch === '"') {
        inString = !inString;
        continue;
      }
      if (inString) continue;
      if (ch === "{") curly += 1;
      else if (ch === "}") curly -= 1;
      else if (ch === "[") square += 1;
      else if (ch === "]") square -= 1;
    }

    if (inString) {
      candidate += '"';
    }

    while (square > 0) {
      candidate += "]";
      square -= 1;
    }
    while (curly > 0) {
      candidate += "}";
      curly -= 1;
    }

    // Clean commas again after auto-closing brackets/braces.
    candidate = candidate.replace(/,\s*([}\]])/g, "$1");

    return JSON.parse(candidate);
  } catch {
    // continue
  }

  throw new Error("Could not parse JSON (possibly truncated or invalid)");
}

async function importLorebookIntoEditor(event) {
  const [file] = event?.target?.files || [];
  if (!file) return;

  try {
    const text = await file.text();
    const parsed = tryParseLooseJson(text);
    const normalized = normalizeLorebookRecord({
      ...parsed,
      name: parsed?.name || parsed?.title || file.name.replace(/\.json$/i, "") || "Imported lorebook",
      description: parsed?.description || "",
    });

    if (els.lorebookEditorView) {
      els.lorebookEditorView.dataset.editingId = "";
    }
    if (els.lorebookNameInput) els.lorebookNameInput.value = normalized.name || "";
    if (els.lorebookDescriptionInput) els.lorebookDescriptionInput.value = normalized.description || "";

    lorebookEntryDraft = (Array.isArray(normalized.entries) ? normalized.entries : [])
      .slice()
      .sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0))
      .map((entry, index) => ({ ...entry, insertionOrder: index }));

    // Fallback parser for unsupported but common lorebook shapes.
    if (!lorebookEntryDraft.length) {
      const scanObjects = [];
      const queue = [parsed];
      const seen = new Set();
      while (queue.length) {
        const current = queue.shift();
        if (!current || typeof current !== "object") continue;
        if (seen.has(current)) continue;
        seen.add(current);

        if (
          current !== parsed
          && (
            Object.prototype.hasOwnProperty.call(current, "content")
            || Object.prototype.hasOwnProperty.call(current, "text")
            || Object.prototype.hasOwnProperty.call(current, "keys")
            || Object.prototype.hasOwnProperty.call(current, "key")
            || Object.prototype.hasOwnProperty.call(current, "keywords")
          )
        ) {
          scanObjects.push(current);
        }

        Object.values(current).forEach((value) => {
          if (value && typeof value === "object") queue.push(value);
        });
      }

      if (scanObjects.length) {
        lorebookEntryDraft = scanObjects
          .map((entry, index) => normalizeLorebookEntry(entry, index))
          .filter((entry) => {
            const hasKeys = Array.isArray(entry.keys) && entry.keys.length > 0;
            const hasSecondary = Array.isArray(entry.secondaryKeys) && entry.secondaryKeys.length > 0;
            const hasTitle = String(entry.title || "").trim().length > 0;
            const hasContent = String(entry.content || "").trim().length > 0;
            return hasContent || hasKeys || hasSecondary || hasTitle || entry.constant;
          })
          .map((entry, index) => ({ ...entry, insertionOrder: index }));
      }
    }

    lorebookEntrySearchQuery = "";
    if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
    closeLorebookEntryEditor();
    renderLorebookEntryDraft();
    setInlineStatus(els.lorebookEditorStatus, `Imported ${normalized.name} (${lorebookEntryDraft.length} entries)`);
    showToast("Lorebook loaded into editor", "success");
  } catch (error) {
    setInlineStatus(els.lorebookEditorStatus, `Could not import lorebook JSON: ${error?.message || "invalid format"}`);
    showToast("Lorebook import failed", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function saveLorebookFromForm() {
  const name = els.lorebookNameInput?.value?.trim() || "";
  const description = els.lorebookDescriptionInput?.value?.trim() || "";

  if (!name) {
    setStatus("Add a lorebook name");
    showToast("Lorebook save failed", "error");
    return;
  }

  if (els.lorebookEntryEditorCard && !els.lorebookEntryEditorCard.classList.contains("hidden")) {
    const ok = saveLorebookEntryFromEditor({ skipClose: true });
    if (!ok) {
      showToast("Fix entry fields before saving lorebook", "error");
      return;
    }
  }

  lorebookEntryDraft = (Array.isArray(lorebookEntryDraft) ? lorebookEntryDraft : [])
    .map((entry, index) => normalizeLorebookEntry({
      ...entry,
      insertion_order: index,
    }, index));

  const editingId = els.lorebookEditorView?.dataset?.editingId || "";
  const existing = editingId ? state.lorebooks.find((book) => book.id === editingId) : null;

  const normalized = normalizeLorebookRecord({
    ...(existing || {}),
    id: editingId || existing?.id || crypto.randomUUID(),
    name,
    description,
    entries: lorebookEntryDraft.map((entry, index) => ({
      id: entry.id || crypto.randomUUID(),
      title: entry.title,
      keys: Array.isArray(entry.keys) ? entry.keys : [],
      secondary_keys: Array.isArray(entry.secondaryKeys) ? entry.secondaryKeys : [],
      content: String(entry.content || "").trim(),
      insertion_order: index,
      priority: Number.isFinite(Number(entry.priority)) ? Number(entry.priority) : 100,
      case_sensitive: Boolean(entry.caseSensitive),
      selective: Boolean(entry.selective),
      selective_logic: entry.selectiveLogic || "AND",
      constant: Boolean(entry.constant),
      probability: Number.isFinite(Number(entry.probability)) ? Number(entry.probability) : 100,
      enabled: entry.enabled !== false,
    })),
  });

  const index = editingId ? state.lorebooks.findIndex((book) => book.id === editingId) : -1;

  if (index >= 0) {
    state.lorebooks[index] = normalized;
    setStatus("Lorebook updated");
  } else {
    state.lorebooks.unshift(normalized);
    setStatus("Lorebook created");
  }

  persistState();
  renderLorebooks();
  renderDrawerLorebooks();
  showLorebookForm(normalized);
  setInlineStatus(els.lorebookEditorStatus, `Lorebook saved (${normalized.entries.length} entries)`);
  showToast("Lorebook saved", "success");
}

async function importLorebooks(event) {
  const files = Array.from(event?.target?.files || []);
  if (!files.length) return;

  try {
    for (const file of files) {
      const text = await file.text();
      const parsed = tryParseLooseJson(text);
      const items = Array.isArray(parsed?.lorebooks)
        ? parsed.lorebooks
        : Array.isArray(parsed)
          ? parsed
          : [parsed];

      items.forEach((item) => {
        const normalized = normalizeLorebookRecord({
          ...item,
          name: item?.name || item?.title || file.name.replace(/\.json$/i, "") || "Imported lorebook",
        });
        state.lorebooks.unshift(normalized);
      });
    }
    persistState();
    renderLorebooks();
    renderDrawerLorebooks();
    setStatus("Lorebook imported");
    showToast("Lorebook imported", "success");
  } catch (error) {
    setStatus(`Could not import lorebook JSON: ${error?.message || "invalid format"}`);
    showToast("Lorebook import failed", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function exportLorebooks() {
  downloadJson("maitavern-lorebooks.json", state.lorebooks);
  setStatus("Lorebooks exported");
  showToast("Lorebooks exported", "success");
}

function exportLorebook(lorebookId) {
  const lorebook = state.lorebooks.find((item) => item.id === lorebookId);
  if (!lorebook) return;
  downloadJson(`${slugify(lorebook.name || "lorebook")}.json`, lorebook);
  setStatus("Lorebook exported");
  showToast("Lorebook exported", "success");
}

function removeLorebook(lorebookId) {
  state.lorebooks = state.lorebooks.filter((item) => item.id !== lorebookId);
  state.chats.forEach((chat) => {
    if (!Array.isArray(chat.lorebookIds)) return;
    chat.lorebookIds = chat.lorebookIds.filter((id) => id !== lorebookId);
  });
  persistState();
  renderLorebooks();
  renderDrawerLorebooks();
  setStatus("Lorebook removed");
  showToast("Lorebook deleted", "success");
}

function renderLorebooks() {
  if (!els.lorebookList) return;

  state.lorebooks = (Array.isArray(state.lorebooks) ? state.lorebooks : []).map((book) => normalizeLorebookRecord(book));

  const query = els.lorebookSearchInput?.value?.trim().toLowerCase() || "";
  const attachedIds = new Set(state.chats.flatMap((chat) => Array.isArray(chat?.lorebookIds) ? chat.lorebookIds : []));

  const lorebooks = state.lorebooks.filter((book) => {
    if (!query) return true;
    const haystack = `${book.name} ${book.description}`.toLowerCase();
    return haystack.includes(query);
  });

  els.lorebookList.innerHTML = "";
  if (!lorebooks.length) {
    const empty = document.createElement("article");
    empty.className = "placeholder-card";
    empty.innerHTML = `<p>No lorebooks yet. Import one or create a new lorebook.</p>`;
    els.lorebookList.appendChild(empty);
    return;
  }

  lorebooks.forEach((book) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-actions">
        <div class="entity-card-header" style="width:100%; margin-bottom:0;">
          <h3>${escapeHtml(book.name || "Untitled lorebook")}</h3>
        </div>
      </div>
      <div class="entity-meta">${escapeHtml(book.description || "No description")}</div>
      <div class="entity-meta">Entries ${book.entries.length} • Default scan ${book.scanDepth} • Default budget ${book.tokenBudget}</div>
      <div class="entity-meta">Attached in ${attachedIds.has(book.id) ? state.chats.filter((chat) => Array.isArray(chat?.lorebookIds) && chat.lorebookIds.includes(book.id)).length : 0} chat(s)</div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="edit" type="button">Edit</button>
        <button class="secondary" data-action="export" type="button">Export</button>
        <button class="danger" data-action="remove" type="button">Delete</button>
      </div>
    `;

    const actions = card.querySelector(".character-inline-actions");
    card.querySelector('[data-action="toggle-actions"]')?.addEventListener("click", () => {
      actions?.classList.toggle("hidden");
    });
    card.querySelector('[data-action="edit"]')?.addEventListener("click", () => showLorebookForm(book));
    card.querySelector('[data-action="export"]')?.addEventListener("click", () => exportLorebook(book.id));
    card.querySelector('[data-action="remove"]')?.addEventListener("click", () => removeLorebook(book.id));

    els.lorebookList.appendChild(card);
  });
}

function normalizePresetPromptList(prompts = []) {
  const toBool = (value, fallback = false) => {
    if (value === undefined || value === null) return fallback;
    if (typeof value === "boolean") return value;
    if (typeof value === "string") {
      const normalized = value.trim().toLowerCase();
      if (["true", "1", "yes", "on"].includes(normalized)) return true;
      if (["false", "0", "no", "off"].includes(normalized)) return false;
    }
    return Boolean(value);
  };

  const toPosition = (value) => {
    if (value === "depth") return 1;
    if (value === "ordered") return 0;
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : 0;
  };

  return prompts
    .map((prompt, index) => {
      const identifier = (prompt?.identifier || `prompt-${index + 1}`).trim();
      return {
        identifier,
        name: prompt?.name || identifier,
        role: prompt?.role || "system",
        content: prompt?.content || "",
        marker: toBool(prompt?.marker, false),
        system_prompt: toBool(prompt?.system_prompt, true),
        enabled: toBool(prompt?.enabled, true),
        injection_position: toPosition(prompt?.injection_position),
        injection_depth: Number.isFinite(Number(prompt?.injection_depth)) ? Number(prompt?.injection_depth) : 4,
        forbid_overrides: toBool(prompt?.forbid_overrides, false),
      };
    })
    .filter((prompt) => prompt.identifier);
}

function normalizePreset(preset) {
  return {
    id: preset.id || crypto.randomUUID(),
    name: preset.name || "Imported preset",
    systemPrompt: preset.systemPrompt || preset.content || defaults.systemPrompt,
    useSystemPrompt: Boolean(preset.useSystemPrompt ?? preset.use_sysprompt ?? true),
    namesBehavior: preset.namesBehavior || preset.names_behavior || "",
    temperature: Number.isFinite(Number(preset.temperature)) ? Number(preset.temperature) : 0.7,
    topP: Number.isFinite(Number(preset.topP ?? preset.top_p)) ? Number(preset.topP ?? preset.top_p) : 1,
    maxTokens: Number.isFinite(Number(preset.maxTokens ?? preset.openai_max_tokens)) ? Number(preset.maxTokens ?? preset.openai_max_tokens) : 300,
    stream: Boolean(preset.stream ?? preset.streaming ?? preset.stream_openai ?? false),
    contextTemplate: preset.contextTemplate || preset.story_string || defaults.presets[0].contextTemplate,
    chatStart: preset.chatStart || preset.chat_start || defaults.presets[0].chatStart,
    exampleSeparator: preset.exampleSeparator || preset.example_separator || defaults.presets[0].exampleSeparator,
    scenarioFormat: preset.scenarioFormat || preset.scenario_format || "{{scenario}}",
    personalityFormat: preset.personalityFormat || preset.personality_format || "{{personality}}",
    prompts: normalizePresetPromptList(Array.isArray(preset.prompts) ? preset.prompts : []),
    promptOrder: Array.isArray(preset.promptOrder) ? preset.promptOrder : Array.isArray(preset.prompt_order) ? preset.prompt_order : [],
    instruct: {
      inputSequence: preset.instruct?.inputSequence || preset.input_sequence || defaults.presets[0].instruct.inputSequence,
      outputSequence: preset.instruct?.outputSequence || preset.output_sequence || defaults.presets[0].instruct.outputSequence,
      systemSequence: preset.instruct?.systemSequence || preset.system_sequence || defaults.presets[0].instruct.systemSequence,
      stopSequence: preset.instruct?.stopSequence || preset.stop_sequence || defaults.presets[0].instruct.stopSequence,
      wrap: Boolean(preset.instruct?.wrap ?? preset.wrap ?? false),
      macro: Boolean(preset.instruct?.macro ?? preset.macro ?? false),
      namesBehavior: preset.instruct?.namesBehavior || preset.names_behavior || "",
      inputSuffix: preset.instruct?.inputSuffix || preset.input_suffix || "",
      outputSuffix: preset.instruct?.outputSuffix || preset.output_suffix || "",
      systemSuffix: preset.instruct?.systemSuffix || preset.system_suffix || "",
      firstInputSequence: preset.instruct?.firstInputSequence || preset.first_input_sequence || "",
      firstOutputSequence: preset.instruct?.firstOutputSequence || preset.first_output_sequence || "",
      lastInputSequence: preset.instruct?.lastInputSequence || preset.last_input_sequence || "",
      lastOutputSequence: preset.instruct?.lastOutputSequence || preset.last_output_sequence || "",
      systemSameAsUser: Boolean(preset.instruct?.systemSameAsUser ?? preset.system_same_as_user ?? false),
      storyStringPrefix: preset.instruct?.storyStringPrefix || preset.story_string_prefix || "",
      storyStringSuffix: preset.instruct?.storyStringSuffix || preset.story_string_suffix || "",
    },
    source: preset.source || "maitavern",
    raw: preset.raw || null,
  };
}

function normalizeEndpoint(endpoint) {
  const trimmed = endpoint.trim().replace(/\/+$/, "");
  if (/\/chat\/completions$/i.test(trimmed)) return trimmed;
  return `${trimmed}/chat/completions`;
}

async function requestChatCompletion({ headers, payload, signal }) {
  const requestUrl = normalizeEndpoint(state.endpoint);
  const response = await fetch(requestUrl, {
    method: "POST",
    headers,
    body: JSON.stringify(payload),
    signal,
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${extractErrorMessage(text) || response.statusText}`);
  }

  return parseApiJsonResponse(text, requestUrl, "chat completion");
}

function shouldRetryWithCompatPayload(data) {
  const choice = Array.isArray(data?.choices) ? data.choices[0] : null;
  return Boolean(choice && choice.message && choice.message.content == null);
}

function extractAssistantText(data) {
  const choice = Array.isArray(data?.choices) ? data.choices[0] : null;
  const messageContent = choice?.message?.content;

  if (typeof messageContent === "string" && messageContent.trim()) {
    return messageContent;
  }

  if (Array.isArray(messageContent)) {
    const joined = messageContent
      .map((item) => {
        if (typeof item === "string") return item;
        if (typeof item?.text === "string") return item.text;
        if (typeof item?.content === "string") return item.content;
        return "";
      })
      .filter(Boolean)
      .join("\n");
    if (joined.trim()) return joined;
  }

  if (typeof choice?.text === "string" && choice.text.trim()) {
    return choice.text;
  }

  if (typeof data?.response === "string" && data.response.trim()) return data.response;
  if (typeof data?.message === "string" && data.message.trim()) return data.message;
  if (typeof data?.content === "string" && data.content.trim()) return data.content;
  if (typeof data?.output_text === "string" && data.output_text.trim()) return data.output_text;

  if (Array.isArray(data?.output)) {
    const outputText = data.output
      .flatMap((item) => item.content || [])
      .map((item) => {
        if (typeof item === "string") return item;
        if (item.type === "output_text") return item.text;
        if (typeof item.text === "string") return item.text;
        return "";
      })
      .filter(Boolean)
      .join("\n");
    if (outputText.trim()) return outputText;
  }

  return "";
}

async function fetchModelsIntoSettings() {
  setInlineStatus(els.fetchModelsStatus, "Fetching models...");
  try {
    const models = await fetchModelsForProvider({
      endpoint: els.endpointInput.value.trim(),
      apiKey: els.apiKeyInput.value.trim(),
      authHeader: els.authHeaderInput.value.trim() || defaults.authHeader,
      authPrefix: els.authPrefixInput.value,
      extraHeadersRaw: els.extraHeadersInput.value.trim(),
    });
    applyFetchedModels(models);
    setInlineStatus(els.fetchModelsStatus, `Loaded ${models.length} model${models.length === 1 ? "" : "s"}`);
    setStatus("Models fetched");
  } catch (error) {
    setInlineStatus(els.fetchModelsStatus, error.message);
    setStatus("Model fetch failed");
  }
}

async function fetchModelsIntoProviderForm() {
  setInlineStatus(els.providerFetchModelsStatus, "Fetching models...");
  try {
    const models = await fetchModelsForProvider({
      endpoint: els.providerUrlInput.value.trim(),
      apiKey: els.providerApiKeyInput.value.trim(),
      authHeader: state.authHeader || defaults.authHeader,
      authPrefix: state.authPrefix,
      extraHeadersRaw: state.extraHeaders,
    });
    applyFetchedModels(models);
    setInlineStatus(els.providerFetchModelsStatus, `Loaded ${models.length} model${models.length === 1 ? "" : "s"}`);
    setStatus("Models fetched");
  } catch (error) {
    setInlineStatus(els.providerFetchModelsStatus, error.message);
    setStatus("Model fetch failed");
  }
}

async function fetchModelsForProvider({ endpoint, apiKey, authHeader, authPrefix, extraHeadersRaw }) {
  if (!endpoint) throw new Error("Add an endpoint first");
  if (!apiKey) throw new Error("Add an API key first");

  const headers = {
    ...parseJsonObject(extraHeadersRaw, "Extra headers"),
  };

  if (apiKey && authHeader) {
    headers[authHeader] = `${authPrefix || ""}${apiKey}`;
  }

  const modelsUrl = normalizeModelsEndpoint(endpoint);
  const response = await fetch(modelsUrl, {
    method: "GET",
    headers,
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${extractErrorMessage(text) || response.statusText}`);
  }

  const data = parseApiJsonResponse(text, modelsUrl, "models");
  const models = extractModels(data);
  if (!models.length) {
    throw new Error("No models returned");
  }
  return models;
}

function normalizeModelsEndpoint(endpoint) {
  const trimmed = endpoint.trim().replace(/\/+$/, "");
  if (/\/models$/i.test(trimmed)) return trimmed;
  if (/\/v1\/chat\/completions$/i.test(trimmed)) {
    return trimmed.replace(/\/chat\/completions$/i, "/models");
  }
  if (/\/chat\/completions$/i.test(trimmed)) {
    return trimmed.replace(/\/chat\/completions$/i, "/models");
  }
  if (/\/v1$/i.test(trimmed)) return `${trimmed}/models`;
  return `${trimmed}/models`;
}

function extractModels(data) {
  const list = Array.isArray(data?.data)
    ? data.data.map((item) => item?.id || item?.name).filter(Boolean)
    : Array.isArray(data?.models)
      ? data.models.map((item) => typeof item === "string" ? item : item?.id || item?.name).filter(Boolean)
      : [];

  return [...new Set(list.map((item) => String(item).trim()).filter(Boolean))];
}

function applyFetchedModels(models) {
  state.models = models;
  state.selectedModel = models.includes(state.selectedModel) ? state.selectedModel : models[0];
  state.hasFetchedModels = true;
  els.modelsInput.value = models.join("\n");
  persistState();
  renderModelOptions();
}

async function maybeAutoFetchModels(force = false) {
  if (!force && state.hasFetchedModels) return;
  try {
    const models = await fetchModelsForProvider({
      endpoint: state.endpoint,
      apiKey: state.apiKey,
      authHeader: state.authHeader || defaults.authHeader,
      authPrefix: state.authPrefix,
      extraHeadersRaw: state.extraHeaders,
    });
    applyFetchedModels(models);
    setInlineStatus(els.fetchModelsStatus, `Loaded ${models.length} model${models.length === 1 ? "" : "s"}`);
    setStatus("Models fetched");
  } catch (error) {
    setInlineStatus(els.fetchModelsStatus, error.message);
  }
}

function setInlineStatus(element, text) {
  if (!element) return;
  element.textContent = text;
}

function setDataStatus(text) {
  if (!els.dataStatus) return;
  els.dataStatus.textContent = text;
}

function parseJsonObject(raw, label) {
  if (!raw) return {};
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error(`${label} must be valid JSON`);
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error(`${label} must be a JSON object`);
  }
  return parsed;
}

async function pickAndLoadUserFile() {
  const { parsed } = await readUserFileFromDisk();
  const importedState = parsed?.state && typeof parsed.state === "object" ? parsed.state : parsed;
  state = {
    ...structuredClone(defaults),
    ...(importedState && typeof importedState === "object" ? importedState : {}),
  };
}

function extractErrorMessage(text) {
  const normalized = String(text || "").trim();
  if (!normalized) return "";
  if (/^<!doctype html/i.test(normalized) || /^<html/i.test(normalized)) {
    return "The server returned HTML instead of JSON. Check the endpoint URL.";
  }
  try {
    const parsed = JSON.parse(normalized);
    return parsed?.error?.message || parsed?.message || normalized;
  } catch {
    return normalized;
  }
}

function parseApiJsonResponse(text, requestUrl, label) {
  const normalized = String(text || "").trim();
  if (!normalized) {
    throw new Error(`Empty ${label} response from ${requestUrl}`);
  }
  if (/^<!doctype html/i.test(normalized) || /^<html/i.test(normalized) || normalized.startsWith("<")) {
    throw new Error(`The ${label} endpoint returned HTML instead of JSON. Check the URL: ${requestUrl}`);
  }
  try {
    return JSON.parse(normalized);
  } catch {
    throw new Error(`The ${label} endpoint returned non-JSON data. Check the URL: ${requestUrl}`);
  }
}

function renderModelOptions() {
  els.modelSelect.innerHTML = "";
  for (const model of state.models) {
    const option = document.createElement("option");
    option.value = model;
    option.textContent = model;
    option.selected = model === state.selectedModel;
    els.modelSelect.appendChild(option);
  }
}

function renderMessages() {
  updateComposerOffset();
  const previousTop = els.messages.scrollTop;
  const shouldStickToBottom = Math.abs((els.messages.scrollHeight - els.messages.clientHeight) - previousTop) < 80;
  els.messages.innerHTML = "";
  for (let index = 0; index < state.messages.length; index += 1) {
    const msg = state.messages[index];
    const row = els.messageTemplate.content.firstElementChild.cloneNode(true);
    row.classList.add(msg.role);
    if (msg.pending) row.classList.add("pending");
    const avatar = row.querySelector(".message-avatar");
    const bubble = row.querySelector(".message");
    bubble.classList.add(msg.role);
    row.querySelector(".message-role").textContent = getMessageDisplayRole(msg);

    const menuTrigger = row.querySelector(".message-menu-trigger");
    menuTrigger?.addEventListener("click", () => openMessageActionPopup(index));

    const thinkingChip = row.querySelector(".thinking-chip");
    const thinkingSummary = row.querySelector(".thinking-summary");
    const activeProvider = getActiveProvider();
    if (msg.thinkingSummary && activeProvider?.showThinkingSummaries !== false) {
      thinkingChip?.classList.remove("hidden");
      thinkingSummary.textContent = msg.thinkingSummary;
      thinkingChip?.addEventListener("click", () => thinkingSummary.classList.toggle("hidden"));
    }

    const activeContent = getMessageSwipeContent(msg);
    row.querySelector(".message-content").innerHTML = formatMessageContent(activeContent);

    const swipeControls = row.querySelector(".message-swipe-controls");
    if (swipeControls && Array.isArray(msg.swipes) && msg.swipes.length > 1) {
      const swipeId = clampSwipeIndex(msg.swipe_id, msg.swipes.length);
      msg.swipe_id = swipeId;
      swipeControls.classList.remove("hidden");
      swipeControls.querySelector(".message-swipe-counter").textContent = `${swipeId + 1}/${msg.swipes.length}`;
      swipeControls.querySelector(".message-swipe-left").addEventListener("click", () => setMessageSwipeIndex(index, swipeId - 1));
      swipeControls.querySelector(".message-swipe-right").addEventListener("click", () => setMessageSwipeIndex(index, swipeId + 1));
    }

    const avatarUrl = getMessageAvatar(msg);
    setElementBackgroundImage(avatar, avatarUrl);
    els.messages.appendChild(row);
  }

  if (shouldStickToBottom) {
    els.messages.scrollTop = els.messages.scrollHeight;
  }
  updateChatUiVisibility();
}

function renderCharacters() {
  const query = els.characterSearchInput.value.trim().toLowerCase();
  const characters = state.characters.filter((character) => {
    if (!query) return true;
    return `${character.name} ${character.description || ""} ${character.personality || ""} ${character.scenario || ""}`.toLowerCase().includes(query);
  });

  els.characterList.innerHTML = "";
  characters.forEach((character) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-actions">
        <div class="character-card-avatar"></div>
        <div class="entity-card-header">
          <h3>${escapeHtml(character.name || "Untitled character")}</h3>
          <span class="entity-meta">SillyTavern JSON compatible</span>
        </div>
      </div>
      <div class="entity-meta">${escapeHtml(character.scenario || character.description || "No scenario yet.")}</div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="edit">Edit</button>
        <button class="secondary" data-action="new-chat">New chat</button>
      </div>
    `;
    setElementBackgroundImage(card.querySelector(".character-card-avatar"), getCharacterAvatarUrl(character));
    const actions = card.querySelector(".character-inline-actions");
    card.querySelector('[data-action="toggle-actions"]').addEventListener("click", () => actions.classList.toggle("hidden"));
    card.querySelector('[data-action="edit"]').addEventListener("click", () => editCharacter(character.id));
    card.querySelector('[data-action="new-chat"]').addEventListener("click", () => createCharacterChat(character.id));
    els.characterList.appendChild(card);
  });
}

function createCharacter() {
  showCharacterEditor();
  switchView("characterStudioView");
  setStatus("Character studio");
}

function createCharacterChat(characterId) {
  const character = state.characters.find((item) => item.id === characterId);
  if (!character) {
    showToast("Character not found", "error");
    return;
  }
  state.messages = [];
  const chat = createChatRecord(character.name || "New chat");
  chat.characterId = character.id;
  ensureGreetingMessageForChat(chat);
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title;
  state.messages = Array.isArray(chat.messages) ? chat.messages.map((message) => ({ ...message })) : [];
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
  renderDrawerState();
  switchView("chatsView");
  showToast("New chat created", "success");
}

async function importCharacters(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const normalized = normalizeImportedCharacter(parsed, file.name);
    showCharacterEditor(normalized);
    switchView("characterStudioView");
    els.characterEditorCard.dataset.importedData = JSON.stringify(normalized.data || parsed);
    setStatus("Character imported into studio");
    showToast("Character imported into studio", "success");
  } catch (error) {
    setStatus("Could not import character JSON");
    showToast(error?.message || "Character import failed", "error");
  } finally {
    event.target.value = "";
  }
}

function editCharacter(characterId) {
  const character = state.characters.find((item) => item.id === characterId);
  if (!character) return;
  showCharacterEditor(character);
  switchView("characterStudioView");
  setStatus(`Editing ${character.name || "character"}`);
}

async function importCharacterFromAvatar(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const imported = await importCharacterCardFromImage(file);
    showCharacterEditor(imported);
    switchView("characterStudioView");
    els.characterEditorCard.dataset.importedData = JSON.stringify(imported.data || {});
    setStatus("Character imported into studio");
    showToast("Character imported into studio", "success");
  } catch (error) {
    setStatus("Could not import avatar image");
    showToast(error?.message || "Avatar import failed", "error");
  } finally {
    event.target.value = "";
  }
}

function showProviderForm(provider = null) {
  els.providerFormCard.classList.remove("hidden");
  els.providerFormCard.dataset.editingId = provider?.id || "";
  els.providerNameInput.value = provider?.name || "";
  els.providerUrlInput.value = provider?.baseUrl || "";
  els.providerApiKeyInput.value = provider?.apiKey || "";
  if (els.providerThinkStartInput) els.providerThinkStartInput.value = provider?.thinkingStartTag || "<thinking>";
  if (els.providerThinkEndInput) els.providerThinkEndInput.value = provider?.thinkingEndTag || "</thinking>";
  if (els.providerShowThinkingInput) els.providerShowThinkingInput.checked = provider?.showThinkingSummaries !== false;
  if (els.providerIncludeThinkingInput) els.providerIncludeThinkingInput.checked = Boolean(provider?.includeThinkingInPrompt);
  els.providerNameInput.focus();
}

function hideProviderForm() {
  els.providerFormCard.classList.add("hidden");
  delete els.providerFormCard.dataset.editingId;
  els.providerNameInput.value = "";
  els.providerUrlInput.value = "";
  els.providerApiKeyInput.value = "";
  if (els.providerThinkStartInput) els.providerThinkStartInput.value = "<thinking>";
  if (els.providerThinkEndInput) els.providerThinkEndInput.value = "</thinking>";
  if (els.providerShowThinkingInput) els.providerShowThinkingInput.checked = true;
  if (els.providerIncludeThinkingInput) els.providerIncludeThinkingInput.checked = false;
  setInlineStatus(els.providerFetchModelsStatus, "");
}

function saveProviderFromForm() {
  const name = els.providerNameInput.value.trim();
  const baseUrl = els.providerUrlInput.value.trim();
  const apiKey = els.providerApiKeyInput.value.trim();
  if (!name || !baseUrl || !apiKey) {
    setStatus("Fill in provider name, base URL, and API key");
    showToast("Provider save failed", "error");
    return;
  }

  const editingId = els.providerFormCard.dataset.editingId;
  const existingIndex = editingId
    ? state.providers.findIndex((provider) => provider.id === editingId)
    : state.providers.findIndex((provider) => provider.name.toLowerCase() === name.toLowerCase());
  const provider = {
    id: existingIndex >= 0 ? state.providers[existingIndex].id : crypto.randomUUID(),
    name,
    baseUrl,
    apiKey,
    thinkingStartTag: els.providerThinkStartInput?.value?.trim() || "<thinking>",
    thinkingEndTag: els.providerThinkEndInput?.value?.trim() || "</thinking>",
    showThinkingSummaries: Boolean(els.providerShowThinkingInput?.checked ?? true),
    includeThinkingInPrompt: Boolean(els.providerIncludeThinkingInput?.checked ?? false),
  };

  if (existingIndex >= 0) {
    state.providers[existingIndex] = provider;
  } else {
    state.providers.unshift(provider);
  }

  state.activeProviderId = provider.id;
  state.endpoint = provider.baseUrl;
  state.apiKey = provider.apiKey;
  persistState();
  renderProviders();
  renderDrawerState();
  syncInputsFromState();
  hideProviderForm();
  setStatus(existingIndex >= 0 ? "Provider updated" : "Provider saved");
  showToast(existingIndex >= 0 ? "Provider updated!" : "Provider created!", "success");
}

async function importProviders(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const providers = Array.isArray(parsed) ? parsed : [parsed];
    providers.forEach((provider, index) => {
      const importedProvider = {
        id: crypto.randomUUID(),
        name: provider.name || "Imported provider",
        baseUrl: provider.baseUrl || provider.endpoint || "",
        apiKey: provider.apiKey || "",
        thinkingStartTag: provider.thinkingStartTag || "<thinking>",
        thinkingEndTag: provider.thinkingEndTag || "</thinking>",
        showThinkingSummaries: provider.showThinkingSummaries !== false,
        includeThinkingInPrompt: Boolean(provider.includeThinkingInPrompt),
      };
      state.providers.unshift(importedProvider);
      if (index === 0 && !state.activeProviderId) {
        state.activeProviderId = importedProvider.id;
      }
    });
    persistState();
    renderProviders();
    renderDrawerState();
    setStatus("Provider imported");
  } catch {
    setStatus("Could not import provider JSON");
  } finally {
    event.target.value = "";
  }
}

function exportProviders() {
  downloadJson("maitavern-providers.json", state.providers.map(({ name, baseUrl, apiKey }) => ({ name, baseUrl, apiKey })));
  setStatus("Providers exported");
}

function exportSingleProvider(providerId) {
  const provider = state.providers.find((item) => item.id === providerId);
  if (!provider) return;
  downloadJson(`${slugify(provider.name || "provider")}.json`, {
    name: provider.name,
    baseUrl: provider.baseUrl,
    apiKey: provider.apiKey,
  });
  setStatus("Provider exported");
  showToast("Provider exported", "success");
}

function renderProviders() {
  const query = els.providerSearchInput.value.trim().toLowerCase();
  const providers = state.providers.filter((provider) => {
    if (!query) return true;
    return `${provider.name} ${provider.baseUrl}`.toLowerCase().includes(query);
  });

  els.providerList.innerHTML = "";
  providers.forEach((provider) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-provider-actions">
        <div>
          <h3>${escapeHtml(provider.name)}</h3>
          <div class="entity-meta">OpenAI-compatible</div>
        </div>
      </div>
      <div class="entity-meta">${escapeHtml(provider.baseUrl)}</div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="use">Use</button>
        <button class="secondary" data-action="edit">Edit</button>
        <button class="secondary" data-action="export">Export</button>
        <button class="secondary" data-action="remove">Delete</button>
      </div>
    `;
    const actions = card.querySelector(".character-inline-actions");
    card.querySelector('[data-action="toggle-provider-actions"]').addEventListener("click", () => actions.classList.toggle("hidden"));
    card.querySelector('[data-action="use"]').addEventListener("click", () => useProvider(provider.id));
    card.querySelector('[data-action="edit"]').addEventListener("click", () => showProviderForm(provider));
    card.querySelector('[data-action="export"]').addEventListener("click", () => exportSingleProvider(provider.id));
    card.querySelector('[data-action="remove"]').addEventListener("click", () => removeProvider(provider.id));
    els.providerList.appendChild(card);
  });

  renderDrawerProviders();
}

function useProvider(providerId) {
  const provider = state.providers.find((item) => item.id === providerId);
  if (!provider) return;
  state.activeProviderId = provider.id;
  state.endpoint = provider.baseUrl;
  state.apiKey = provider.apiKey;
  syncInputsFromState();
  persistState();
  renderDrawerState();
  els.providerPicker.classList.add("hidden");
  setStatus(`Using provider ${provider.name}`);
}

function removeProvider(providerId) {
  state.providers = state.providers.filter((item) => item.id !== providerId);
  if (state.activeProviderId === providerId) {
    state.activeProviderId = "";
  }
  persistState();
  renderProviders();
  renderDrawerState();
  setStatus("Provider removed");
  showToast("Provider deleted", "success");
}

function renderDrawerProviders() {
  if (!els.drawerProviderList) return;
  els.drawerProviderList.innerHTML = "";
  state.providers.forEach((provider) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `drawer-choice${provider.id === state.activeProviderId ? " active" : ""}`;
    button.innerHTML = `
      <strong>${escapeHtml(provider.name)}</strong>
      <small>${escapeHtml(provider.baseUrl)}</small>
    `;
    button.addEventListener("click", () => {
      useProvider(provider.id);
      showDrawerProviderEditor(provider.id);
    });
    els.drawerProviderList.appendChild(button);
  });
}

function getActiveChatLorebookIds() {
  const chat = getCurrentChat();
  return Array.isArray(chat?.lorebookIds)
    ? [...new Set(chat.lorebookIds.map((item) => String(item || "").trim()).filter(Boolean))]
    : [];
}

function getActiveChatLorebookSettings() {
  const chat = getCurrentChat();
  const source = chat?.lorebookSettings && typeof chat.lorebookSettings === "object"
    ? chat.lorebookSettings
    : defaults.lorebookSettings;

  return {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(source?.scanDepth, 10) || defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(source?.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: source?.recursiveScanning !== false,
  };
}

function saveActiveChatLorebookIds(nextIds = []) {
  const chat = getCurrentChat();
  if (!chat) {
    setInlineStatus(els.drawerLorebookStatus, "Open a chat first");
    return;
  }

  chat.lorebookIds = [...new Set(nextIds.map((item) => String(item || "").trim()).filter(Boolean))];
  syncCurrentChatFromState();
  persistState();
  renderDrawerLorebooks();
}

function renderDrawerLorebooks() {
  if (!els.drawerLorebookList) return;

  const chat = getCurrentChat();
  const isChatOpen = Boolean(chat?.id);
  const query = String(els.drawerLorebookSearchInput?.value || "").trim().toLowerCase();
  const selectedIds = new Set(getActiveChatLorebookIds());
  const visibleLorebooks = state.lorebooks.filter((book) => {
    if (!query) return true;
    return `${book.name || ""} ${book.description || ""}`.toLowerCase().includes(query);
  });

  els.drawerLorebookList.innerHTML = "";

  if (!isChatOpen) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">Open a chat to attach lorebooks.</div>`;
    els.drawerLorebookList.appendChild(empty);
  } else if (!state.lorebooks.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No lorebooks in library yet.</div>`;
    els.drawerLorebookList.appendChild(empty);
  } else if (!visibleLorebooks.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No lorebooks match your search.</div>`;
    els.drawerLorebookList.appendChild(empty);
  } else {
    visibleLorebooks.forEach((book) => {
      const button = document.createElement("button");
      button.type = "button";
      const active = selectedIds.has(book.id);
      button.className = `drawer-choice${active ? " active" : ""}`;
      button.innerHTML = `
        <strong>${escapeHtml(book.name || "Untitled lorebook")}</strong>
        <small>${book.entries?.length || 0} entries${active ? " • attached" : ""}</small>
      `;
      button.addEventListener("click", () => {
        if (!getCurrentChat()) {
          setInlineStatus(els.drawerLorebookStatus, "Open a chat first");
          return;
        }
        const next = new Set(getActiveChatLorebookIds());
        if (next.has(book.id)) next.delete(book.id);
        else next.add(book.id);
        saveActiveChatLorebookIds([...next]);
        setInlineStatus(els.drawerLorebookStatus, next.has(book.id) ? `Attached ${book.name}` : `Detached ${book.name}`);
      });
      els.drawerLorebookList.appendChild(button);
    });
  }

  const selectedNames = state.lorebooks
    .filter((book) => selectedIds.has(book.id))
    .map((book) => book.name)
    .filter(Boolean);
  if (els.currentLorebookLabel) {
    els.currentLorebookLabel.textContent = `Lorebooks in this chat: ${selectedNames.length ? selectedNames.join(", ") : "none"}`;
  }

  const chatSettings = getActiveChatLorebookSettings();
  if (els.drawerLorebookScanDepthInput) els.drawerLorebookScanDepthInput.value = String(chatSettings.scanDepth);
  if (els.drawerLorebookTokenBudgetInput) els.drawerLorebookTokenBudgetInput.value = String(chatSettings.tokenBudget);
  if (els.drawerLorebookRecursiveInput) els.drawerLorebookRecursiveInput.checked = chatSettings.recursiveScanning;
}

function saveDrawerLorebookSettings() {
  const chat = getCurrentChat();
  if (!chat) {
    setInlineStatus(els.drawerLorebookStatus, "Open a chat first");
    return;
  }

  chat.lorebookSettings = {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(els.drawerLorebookScanDepthInput?.value || "", 10) || defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(els.drawerLorebookTokenBudgetInput?.value || "", 10) || defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: Boolean(els.drawerLorebookRecursiveInput?.checked ?? true),
  };

  syncCurrentChatFromState();
  persistState();
  renderDrawerLorebooks();
  setInlineStatus(els.drawerLorebookStatus, "Lore settings saved for current chat");
}

function renderPresets() {
  if (!els.presetList) return;
  els.presetList.innerHTML = "";
  state.presets.forEach((preset) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="entity-card-header">
        <h3>${escapeHtml(preset.name)}</h3>
        <span class="entity-meta">${escapeHtml((preset.source || "maitavern").toUpperCase())}</span>
      </div>
      <div class="entity-meta">Temp ${Number(preset.temperature ?? 0.7).toFixed(2)} • Top P ${Number(preset.topP ?? 1).toFixed(2)} • Max ${escapeHtml(preset.maxTokens ?? 300)} • Stream ${preset.stream ? "On" : "Off"}</div>
      <div class="entity-meta">System prompt ${preset.useSystemPrompt !== false ? "On" : "Off"} • Context ${preset.contextTemplate ? "loaded" : "none"} • Instruct ${shouldApplyInstructPreset(preset) ? "loaded" : "none"}</div>
      <div class="entity-meta">Prompt blocks ${Array.isArray(preset.prompts) ? preset.prompts.length : 0} • Prompt order ${Array.isArray(preset.promptOrder) && preset.promptOrder.length ? "loaded" : "none"}</div>
      <div class="entity-meta">${escapeHtml((preset.systemPrompt || "No system prompt").slice(0, 140))}</div>
      <div class="entity-card-actions">
        <button class="secondary" data-action="use">Use</button>
        <button class="secondary" data-action="edit">Edit</button>
        <button class="secondary" data-action="export">Export</button>
        <button class="secondary" data-action="remove">Remove</button>
      </div>
    `;
    card.querySelector('[data-action="use"]').addEventListener("click", () => usePreset(preset.id));
    card.querySelector('[data-action="edit"]').addEventListener("click", () => showPresetForm(preset));
    card.querySelector('[data-action="export"]').addEventListener("click", () => exportPreset(preset.id));
    card.querySelector('[data-action="remove"]').addEventListener("click", () => removePreset(preset.id));
    els.presetList.appendChild(card);
  });

  renderDrawerPresets();
}

function showPresetForm(preset = null) {
  const normalized = preset ? normalizePreset(preset) : normalizePreset({ name: "", systemPrompt: "", source: "maitavern" });
  presetEditorDraft = {
    id: normalized.id,
    prompts: getPromptDraftListFromPreset(normalized),
  };

  els.presetFormCard.classList.remove("hidden");
  els.presetFormCard.dataset.editingId = preset?.id || "";
  els.presetNameInput.value = preset?.name || "";
  els.presetTemperatureInput.value = normalized.temperature;
  els.presetTopPInput.value = normalized.topP;
  els.presetMaxTokensInput.value = normalized.maxTokens;
  els.presetStreamInput.checked = Boolean(normalized.stream);
  els.presetSystemPromptInput.value = normalized.systemPrompt || "";
  els.presetContextTemplateInput.value = normalized.contextTemplate || "";
  els.presetChatStartInput.value = normalized.chatStart || "";
  els.presetExampleSeparatorInput.value = normalized.exampleSeparator || "";
  els.presetInstructUserInput.value = normalized.instruct.inputSequence || "";
  els.presetInstructAssistantInput.value = normalized.instruct.outputSequence || "";
  els.presetInstructSystemInput.value = normalized.instruct.systemSequence || "";
  els.presetStopSequenceInput.value = normalized.instruct.stopSequence || "";
  if (els.presetUseSystemPromptInput) els.presetUseSystemPromptInput.checked = normalized.useSystemPrompt !== false;
  if (els.presetNamesBehaviorInput) els.presetNamesBehaviorInput.value = normalized.namesBehavior || normalized.instruct?.namesBehavior || "";
  if (els.presetScenarioFormatInput) els.presetScenarioFormatInput.value = normalized.scenarioFormat || "{{scenario}}";
  if (els.presetPersonalityFormatInput) els.presetPersonalityFormatInput.value = normalized.personalityFormat || "{{personality}}";

  renderPresetPromptList();
  hidePresetPromptEditor();
  els.presetNameInput.focus();
}

function hidePresetForm() {
  els.presetFormCard.classList.add("hidden");
  delete els.presetFormCard.dataset.editingId;
  presetEditorDraft = null;
  hidePresetPromptEditor();
  els.presetNameInput.value = "";
  els.presetTemperatureInput.value = "0.7";
  els.presetTopPInput.value = "1";
  els.presetMaxTokensInput.value = "300";
  els.presetStreamInput.checked = Boolean(defaults.presets[0].stream);
  els.presetSystemPromptInput.value = "";
  els.presetContextTemplateInput.value = defaults.presets[0].contextTemplate;
  els.presetChatStartInput.value = defaults.presets[0].chatStart;
  els.presetExampleSeparatorInput.value = defaults.presets[0].exampleSeparator;
  els.presetInstructUserInput.value = defaults.presets[0].instruct.inputSequence;
  els.presetInstructAssistantInput.value = defaults.presets[0].instruct.outputSequence;
  els.presetInstructSystemInput.value = defaults.presets[0].instruct.systemSequence;
  els.presetStopSequenceInput.value = defaults.presets[0].instruct.stopSequence;
  if (els.presetUseSystemPromptInput) els.presetUseSystemPromptInput.checked = true;
  if (els.presetNamesBehaviorInput) els.presetNamesBehaviorInput.value = "";
  if (els.presetScenarioFormatInput) els.presetScenarioFormatInput.value = "{{scenario}}";
  if (els.presetPersonalityFormatInput) els.presetPersonalityFormatInput.value = "{{personality}}";
  if (els.presetPromptList) els.presetPromptList.innerHTML = "";
}

function getPromptDraftListFromPreset(preset) {
  return getOrderedPromptEntries(preset).map((prompt) => ({
    ...prompt,
    enabled: prompt._orderEnabled !== false && prompt.enabled !== false,
  }));
}

function renderPresetPromptList() {
  if (!els.presetPromptList) return;
  const prompts = Array.isArray(presetEditorDraft?.prompts) ? presetEditorDraft.prompts : [];
  els.presetPromptList.innerHTML = "";

  if (!prompts.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No prompt blocks yet. Add one to create toggles/injections like SillyTavern.</div>`;
    els.presetPromptList.appendChild(empty);
    return;
  }

  prompts.forEach((prompt, index) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-actions">
        <div class="entity-card-header" style="width:100%; margin-bottom:0;">
          <h3>${escapeHtml(prompt.name || prompt.identifier || `Prompt ${index + 1}`)}</h3>
          <button class="secondary" data-action="enable-toggle" type="button">${prompt.enabled !== false ? "On" : "Off"}</button>
        </div>
      </div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="edit" type="button">Edit</button>
        <button class="danger" data-action="remove" type="button">Delete</button>
      </div>
    `;

    const actions = card.querySelector(".character-inline-actions");

    card.querySelector('[data-action="toggle-actions"]').addEventListener("click", (event) => {
      if (event.target.closest('[data-action="enable-toggle"]')) return;
      actions.classList.toggle("hidden");
    });

    card.querySelector('[data-action="enable-toggle"]').addEventListener("click", (event) => {
      event.stopPropagation();
      prompt.enabled = prompt.enabled === false;
      renderPresetPromptList();
    });

    card.querySelector('[data-action="edit"]').addEventListener("click", () => showPresetPromptEditor(index));

    card.querySelector('[data-action="remove"]').addEventListener("click", () => {
      presetEditorDraft.prompts.splice(index, 1);
      hidePresetPromptEditor();
      renderPresetPromptList();
    });

    els.presetPromptList.appendChild(card);
  });
}

function createPresetPromptDraft() {
  if (!presetEditorDraft) presetEditorDraft = { prompts: [] };
  const nextIndex = (presetEditorDraft.prompts?.length || 0) + 1;
  const prompt = {
    identifier: `custom-${Date.now()}`,
    name: `Custom prompt ${nextIndex}`,
    role: "system",
    content: "",
    marker: false,
    system_prompt: true,
    enabled: true,
    injection_position: 0,
    injection_depth: 4,
    forbid_overrides: false,
  };
  presetEditorDraft.prompts.push(prompt);
  renderPresetPromptList();
  showPresetPromptEditor(presetEditorDraft.prompts.length - 1);
}

function showPresetPromptEditor(index) {
  if (!els.presetPromptEditor || !presetEditorDraft?.prompts?.[index]) return;
  const prompt = presetEditorDraft.prompts[index];
  els.presetPromptModalBackdrop?.classList.remove("hidden");
  els.presetPromptEditor.classList.remove("hidden");
  els.presetPromptEditor.dataset.promptIndex = String(index);
  els.presetPromptIdentifierInput.value = prompt.identifier || "";
  els.presetPromptNameInput.value = prompt.name || "";
  els.presetPromptRoleInput.value = prompt.role || "system";
  els.presetPromptInjectionPositionInput.value = String(Number(prompt.injection_position ?? 0));
  els.presetPromptInjectionDepthInput.value = String(Number(prompt.injection_depth ?? 4));
  els.presetPromptEnabledInput.checked = prompt.enabled !== false;
  els.presetPromptSystemPromptInput.checked = prompt.system_prompt !== false;
  els.presetPromptMarkerInput.checked = Boolean(prompt.marker);
  els.presetPromptForbidOverridesInput.checked = Boolean(prompt.forbid_overrides);
  els.presetPromptContentInput.value = prompt.content || "";
}

function hidePresetPromptEditor() {
  if (!els.presetPromptEditor) return;
  els.presetPromptModalBackdrop?.classList.add("hidden");
  els.presetPromptEditor.classList.add("hidden");
  delete els.presetPromptEditor.dataset.promptIndex;
}

function savePresetPromptDraft() {
  const index = Number(els.presetPromptEditor?.dataset.promptIndex);
  if (!Number.isFinite(index) || !presetEditorDraft?.prompts?.[index]) return;

  const identifier = els.presetPromptIdentifierInput.value.trim();
  const name = els.presetPromptNameInput.value.trim();
  if (!identifier) {
    showToast("Prompt identifier is required", "error");
    return;
  }

  presetEditorDraft.prompts[index] = {
    ...presetEditorDraft.prompts[index],
    identifier,
    name: name || identifier,
    role: els.presetPromptRoleInput.value || "system",
    injection_position: Number.parseInt(els.presetPromptInjectionPositionInput.value, 10) || 0,
    injection_depth: Number.parseInt(els.presetPromptInjectionDepthInput.value, 10) || 0,
    enabled: Boolean(els.presetPromptEnabledInput.checked),
    system_prompt: Boolean(els.presetPromptSystemPromptInput.checked),
    marker: Boolean(els.presetPromptMarkerInput.checked),
    forbid_overrides: Boolean(els.presetPromptForbidOverridesInput.checked),
    content: els.presetPromptContentInput.value,
  };

  renderPresetPromptList();
  hidePresetPromptEditor();
}

function deletePresetPromptDraft() {
  const index = Number(els.presetPromptEditor?.dataset.promptIndex);
  if (!Number.isFinite(index) || !presetEditorDraft?.prompts?.[index]) return;
  presetEditorDraft.prompts.splice(index, 1);
  hidePresetPromptEditor();
  renderPresetPromptList();
}

function savePresetFromForm() {
  const name = els.presetNameInput.value.trim();
  const systemPrompt = els.presetSystemPromptInput.value.trim() || defaults.systemPrompt;
  const temperature = Number.parseFloat(els.presetTemperatureInput.value);
  const topP = Number.parseFloat(els.presetTopPInput.value);
  const maxTokens = Number.parseInt(els.presetMaxTokensInput.value, 10);
  const stream = Boolean(els.presetStreamInput.checked);

  if (!name) {
    setStatus("Add a preset name");
    return;
  }

  if (!Number.isFinite(temperature) || temperature < 0 || temperature > 2) {
    setStatus("Temperature must be between 0 and 2");
    return;
  }

  if (!Number.isFinite(topP) || topP < 0 || topP > 1) {
    setStatus("Top P must be between 0 and 1");
    return;
  }

  if (!Number.isFinite(maxTokens) || maxTokens < 1) {
    setStatus("Max tokens must be greater than 0");
    return;
  }

  const promptDraftList = normalizePresetPromptList(presetEditorDraft?.prompts || []);
  const promptOrder = [{
    character_id: 100001,
    order: promptDraftList.map((prompt) => ({
      identifier: prompt.identifier,
      enabled: prompt.enabled !== false,
    })),
  }];

  const presetData = normalizePreset({
    id: els.presetFormCard.dataset.editingId || crypto.randomUUID(),
    name,
    systemPrompt,
    useSystemPrompt: Boolean(els.presetUseSystemPromptInput?.checked ?? true),
    names_behavior: els.presetNamesBehaviorInput?.value?.trim() || "",
    temperature,
    topP,
    maxTokens,
    stream,
    contextTemplate: els.presetContextTemplateInput.value.trim(),
    chatStart: els.presetChatStartInput.value,
    exampleSeparator: els.presetExampleSeparatorInput.value,
    scenarioFormat: els.presetScenarioFormatInput?.value?.trim() || "{{scenario}}",
    personalityFormat: els.presetPersonalityFormatInput?.value?.trim() || "{{personality}}",
    prompts: promptDraftList,
    prompt_order: promptOrder,
    instruct: {
      inputSequence: els.presetInstructUserInput.value,
      outputSequence: els.presetInstructAssistantInput.value,
      systemSequence: els.presetInstructSystemInput.value,
      stopSequence: els.presetStopSequenceInput.value,
      namesBehavior: els.presetNamesBehaviorInput?.value?.trim() || "",
    },
    source: "maitavern",
  });

  const editingId = els.presetFormCard.dataset.editingId;
  const presetIndex = editingId ? state.presets.findIndex((item) => item.id === editingId) : -1;

  if (presetIndex >= 0) {
    state.presets[presetIndex] = presetData;
    setStatus("Preset updated");
  } else {
    state.presets.unshift(presetData);
    state.activePresetId = presetData.id;
    setStatus("Preset created");
  }

  persistState();
  renderPresets();
  renderDrawerState();
  hidePresetForm();
}

async function importPresets(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;

  try {
    for (const file of files) {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const importedPreset = normalizeImportedPreset(parsed, file.name);
      state.presets.unshift(importedPreset);
    }
    persistState();
    renderPresets();
    renderDrawerState();
    setStatus("Preset imported");
  } catch {
    setStatus("Could not import preset JSON");
  } finally {
    event.target.value = "";
  }
}

function exportPresets() {
  downloadJson("maitavern-presets.json", state.presets);
  setStatus("Presets exported");
}

function exportPreset(presetId) {
  const preset = state.presets.find((item) => item.id === presetId);
  if (!preset) return;
  downloadJson(`${slugify(preset.name || "preset")}.json`, preset);
  setStatus("Preset exported");
}

function normalizeImportedPreset(parsed, fileName = "") {
  if (parsed?.chat_completion_source || parsed?.prompts || parsed?.openai_setting_names) {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported OpenAI preset",
      systemPrompt: parsed.prompts?.find((prompt) => prompt.identifier === "main")?.content
        || parsed.prompts?.find((prompt) => prompt.system_prompt && prompt.content)?.content
        || parsed.system_prompt
        || defaults.systemPrompt,
      use_sysprompt: parsed.use_sysprompt,
      temperature: parsed.temperature,
      top_p: parsed.top_p,
      openai_max_tokens: parsed.openai_max_tokens,
      stream: parsed.stream_openai ?? parsed.stream ?? false,
      scenario_format: parsed.scenario_format || parsed.scenarioFormat,
      personality_format: parsed.personality_format || parsed.personalityFormat,
      prompts: parsed.prompts,
      prompt_order: parsed.prompt_order,
      names_behavior: parsed.names_behavior,
      source: "sillytavern-openai",
      raw: parsed,
    });
  }

  if (parsed?.story_string || parsed?.chat_start || parsed?.example_separator || parsed?.trim_sentences !== undefined) {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported context preset",
      contextTemplate: parsed.story_string,
      chat_start: parsed.chat_start,
      example_separator: parsed.example_separator,
      scenario_format: parsed.scenario_format || parsed.scenarioFormat,
      personality_format: parsed.personality_format || parsed.personalityFormat,
      source: "sillytavern-context",
      raw: parsed,
    });
  }

  if (parsed?.input_sequence || parsed?.output_sequence || parsed?.system_sequence || parsed?.names_behavior !== undefined) {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported instruct preset",
      input_sequence: parsed.input_sequence,
      output_sequence: parsed.output_sequence,
      system_sequence: parsed.system_sequence,
      stop_sequence: parsed.stop_sequence,
      wrap: parsed.wrap,
      macro: parsed.macro,
      names_behavior: parsed.names_behavior,
      input_suffix: parsed.input_suffix,
      output_suffix: parsed.output_suffix,
      system_suffix: parsed.system_suffix,
      first_input_sequence: parsed.first_input_sequence,
      first_output_sequence: parsed.first_output_sequence,
      last_input_sequence: parsed.last_input_sequence,
      last_output_sequence: parsed.last_output_sequence,
      system_same_as_user: parsed.system_same_as_user,
      story_string_prefix: parsed.story_string_prefix,
      story_string_suffix: parsed.story_string_suffix,
      source: "sillytavern-instruct",
      raw: parsed,
    });
  }

  if (typeof parsed?.content === "string") {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported system prompt",
      systemPrompt: parsed.content,
      source: "sillytavern-sysprompt",
      raw: parsed,
    });
  }

  return normalizePreset({
    ...parsed,
    name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported preset",
    source: parsed.source || "imported",
    raw: parsed,
  });
}

function usePreset(presetId) {
  const preset = state.presets.find((item) => item.id === presetId);
  if (!preset) return;
  state.activePresetId = preset.id;
  persistState();
  renderDrawerState();
  els.presetPicker.classList.add("hidden");
  setStatus(`Using preset ${preset.name}`);
}

function removePreset(presetId) {
  if (state.presets.length <= 1) {
    setStatus("Keep at least one preset");
    return;
  }

  state.presets = state.presets.filter((item) => item.id !== presetId);
  if (state.activePresetId === presetId) {
    state.activePresetId = state.presets[0]?.id || defaults.activePresetId;
  }
  persistState();
  renderPresets();
  renderDrawerState();
  setStatus("Preset removed");
}

function renderDrawerPresets() {
  if (!els.drawerPresetList) return;
  els.drawerPresetList.innerHTML = "";
  state.presets.forEach((preset) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `drawer-choice${preset.id === state.activePresetId ? " active" : ""}`;
    button.innerHTML = `
      <strong>${escapeHtml(preset.name)}</strong>
      <small>Temp ${Number(preset.temperature ?? 0.7).toFixed(1)}</small>
    `;
    button.addEventListener("click", () => {
      usePreset(preset.id);
      showDrawerPresetEditor(preset.id);
    });
    els.drawerPresetList.appendChild(button);
  });
}

function openChat(chatId) {
  const chat = state.chats.find((item) => item.id === chatId);
  if (!chat) return;
  ensureGreetingMessageForChat(chat);
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title || "New chat";
  state.messages = Array.isArray(chat.messages) ? chat.messages.map((message) => ({ ...message })) : [];
  renderMessages();
  renderDrawerState();
  switchView("chatsView");
  updateChatUiVisibility();
  persistState();
  setStatus(`Opened ${state.currentChatTitle}`);
}

function renderRecentChats() {
  els.recentChats.innerHTML = "";
  state.recentChats.forEach((chat) => {
    const character = state.characters.find((item) => item.id === chat.characterId);
    const card = document.createElement("article");
    card.className = "chat-card";
    card.innerHTML = `
      <div class="character-card-top">
        <div class="character-card-avatar"></div>
        <div>
          <h3>${escapeHtml(chat.title)}</h3>
          <div class="entity-meta">${escapeHtml(chat.preview)}</div>
        </div>
      </div>
    `;
    setElementBackgroundImage(card.querySelector(".character-card-avatar"), getCharacterAvatarUrl(character));
    card.addEventListener("click", () => openChat(chat.id));
    els.recentChats.appendChild(card);
  });
}

function updateChatUiVisibility() {
  const isChatView = state.activeView === "chatsView";
  const hasActiveChat = isChatView && Boolean(state.currentChatId);
  const hasChatList = isChatView && state.recentChats.length > 0;
  els.composer.classList.toggle("hidden", !hasActiveChat);
  els.messages.classList.toggle("hidden", !isChatView || !hasActiveChat);
  els.chatEmptyState.classList.toggle("hidden", !isChatView || hasActiveChat || hasChatList);
  els.recentChats.classList.toggle("hidden", !isChatView || hasActiveChat);
}

function setBusy(isBusy) {
  els.sendBtn.disabled = isBusy;
  els.messageInput.disabled = isBusy;
  els.impersonateSelect.disabled = isBusy;
  els.stopBtn.classList.toggle("hidden", !isBusy);
  if (isBusy) updateComposerOffset();
}

function updateComposerOffset() {
  const composerHeight = els.composer?.offsetHeight || 0;
  document.documentElement.style.setProperty("--composer-offset", `${composerHeight + 24}px`);
}

function showDrawerPresetEditor(presetId) {
  const preset = state.presets.find((item) => item.id === presetId || item.id === state.activePresetId);
  if (!preset || !els.drawerPresetEditor) return;
  els.drawerPresetEditor.classList.remove("hidden");
  els.drawerPresetEditor.dataset.presetId = preset.id;
  if (els.drawerPresetNameInput) els.drawerPresetNameInput.value = preset.name || "";
  if (els.drawerPresetTempInput) els.drawerPresetTempInput.value = String(preset.temperature ?? 0.7);
  if (els.drawerPresetSystemInput) els.drawerPresetSystemInput.value = preset.systemPrompt || "";
}

function hideDrawerPresetEditor() {
  if (!els.drawerPresetEditor) return;
  els.drawerPresetEditor.classList.add("hidden");
  delete els.drawerPresetEditor.dataset.presetId;
}

function saveDrawerPresetEdits() {
  const presetId = els.drawerPresetEditor?.dataset?.presetId || state.activePresetId;
  const presetIndex = state.presets.findIndex((item) => item.id === presetId);
  if (presetIndex < 0) return;
  const current = state.presets[presetIndex];
  const nextTemp = Number.parseFloat(els.drawerPresetTempInput?.value || String(current.temperature ?? 0.7));
  state.presets[presetIndex] = {
    ...current,
    name: els.drawerPresetNameInput?.value?.trim() || current.name,
    temperature: Number.isFinite(nextTemp) ? nextTemp : current.temperature,
    systemPrompt: els.drawerPresetSystemInput?.value?.trim() || current.systemPrompt,
  };
  persistState();
  renderPresets();
  renderDrawerState();
  setStatus("Preset updated");
}

function showDrawerProviderEditor(providerId) {
  const provider = state.providers.find((item) => item.id === providerId || item.id === state.activeProviderId);
  if (!provider || !els.drawerProviderEditor) return;
  els.drawerProviderEditor.classList.remove("hidden");
  els.drawerProviderEditor.dataset.providerId = provider.id;
  if (els.drawerProviderNameInput) els.drawerProviderNameInput.value = provider.name || "";
  if (els.drawerProviderUrlInput) els.drawerProviderUrlInput.value = provider.baseUrl || "";
  if (els.drawerProviderApiKeyInput) els.drawerProviderApiKeyInput.value = provider.apiKey || "";
  if (els.drawerProviderThinkStartInput) els.drawerProviderThinkStartInput.value = provider.thinkingStartTag || "<thinking>";
  if (els.drawerProviderThinkEndInput) els.drawerProviderThinkEndInput.value = provider.thinkingEndTag || "</thinking>";
  if (els.drawerProviderShowThinkingInput) els.drawerProviderShowThinkingInput.checked = provider.showThinkingSummaries !== false;
  if (els.drawerProviderIncludeThinkingInput) els.drawerProviderIncludeThinkingInput.checked = Boolean(provider.includeThinkingInPrompt);
}

function hideDrawerProviderEditor() {
  if (!els.drawerProviderEditor) return;
  els.drawerProviderEditor.classList.add("hidden");
  delete els.drawerProviderEditor.dataset.providerId;
}

function saveDrawerProviderEdits() {
  const providerId = els.drawerProviderEditor?.dataset?.providerId || state.activeProviderId;
  const providerIndex = state.providers.findIndex((item) => item.id === providerId);
  if (providerIndex < 0) return;
  const current = state.providers[providerIndex];
  const updated = {
    ...current,
    name: els.drawerProviderNameInput?.value?.trim() || current.name,
    baseUrl: els.drawerProviderUrlInput?.value?.trim() || current.baseUrl,
    apiKey: els.drawerProviderApiKeyInput?.value?.trim() || current.apiKey,
    thinkingStartTag: els.drawerProviderThinkStartInput?.value?.trim() || current.thinkingStartTag || "<thinking>",
    thinkingEndTag: els.drawerProviderThinkEndInput?.value?.trim() || current.thinkingEndTag || "</thinking>",
    showThinkingSummaries: Boolean(els.drawerProviderShowThinkingInput?.checked ?? true),
    includeThinkingInPrompt: Boolean(els.drawerProviderIncludeThinkingInput?.checked ?? false),
  };
  state.providers[providerIndex] = updated;
  if (updated.id === state.activeProviderId) {
    state.endpoint = updated.baseUrl;
    state.apiKey = updated.apiKey;
    syncInputsFromState();
  }
  persistState();
  renderProviders();
  renderDrawerState();
  setStatus("Provider updated");
}

function openMessageActionPopup(index) {
  messageActionState = { index, mode: "" };
  if (!els.messageActionPopup || !els.messageActionCard) return;
  const trigger = document.querySelectorAll(".message-menu-trigger")[index];
  if (trigger) {
    const rect = trigger.getBoundingClientRect();
    const top = Math.min(window.innerHeight - 260, rect.bottom + 6);
    const left = Math.min(window.innerWidth - 340, rect.left - 220);
    els.messageActionPopup.style.top = `${Math.max(8, top)}px`;
    els.messageActionPopup.style.left = `${Math.max(8, left)}px`;
  }
  els.messageDeleteOptions?.classList.add("hidden");
  els.messageEditFieldWrap?.classList.add("hidden");
  els.messageActionConfirmBtn?.classList.add("hidden");
  els.messageActionPopup.classList.remove("hidden");
}

function closeMessageActionPopup() {
  if (!els.messageActionPopup) return;
  els.messageActionPopup.classList.add("hidden");
  els.messageActionPopup.style.top = "";
  els.messageActionPopup.style.left = "";
  els.messageDeleteOptions?.classList.add("hidden");
  els.messageEditFieldWrap?.classList.add("hidden");
  els.messageActionConfirmBtn?.classList.add("hidden");
  messageActionState = { index: -1, mode: "" };
}

function startMessageEditFlow() {
  const msg = state.messages[messageActionState.index];
  if (!msg) return;
  messageActionState.mode = "edit";
  els.messageEditFieldWrap?.classList.remove("hidden");
  els.messageActionConfirmBtn?.classList.remove("hidden");
  if (els.messageEditInput) els.messageEditInput.value = msg.content || "";
}

function commitMessageEditFlow() {
  if (messageActionState.mode !== "edit") return;
  const msg = state.messages[messageActionState.index];
  if (!msg) return;
  msg.content = els.messageEditInput?.value || msg.content;
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
  closeMessageActionPopup();
}

function startMessageDeleteFlow() {
  const index = messageActionState.index;
  if (index < 0 || index >= state.messages.length) return;
  const below = Math.max(0, state.messages.length - index - 1);
  if (els.messageDeletePromptText) {
    els.messageDeletePromptText.textContent = below
      ? `Delete this message only, or delete this + ${below} message(s) below?`
      : "Delete this message?";
  }
  els.messageDeleteOptions?.classList.remove("hidden");
}

function commitMessageDeleteFlow(deleteBelow = false) {
  const index = messageActionState.index;
  if (index < 0 || index >= state.messages.length) return;
  if (deleteBelow) {
    state.messages = state.messages.slice(0, index);
  } else {
    state.messages.splice(index, 1);
  }
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
  closeMessageActionPopup();
}

async function regenerateFromMessageAction() {
  const index = messageActionState.index;
  if (index < 0 || index >= state.messages.length) return;
  const msg = state.messages[index];
  if (!msg || msg.role !== "assistant") {
    showToast("Regenerate is available on assistant messages", "error");
    return;
  }

  state.messages = state.messages.slice(0, index);
  const assistantMessage = { role: "assistant", content: "", pending: true };
  state.messages.push(assistantMessage);
  const requestMessages = state.messages.slice(0, -1);

  syncCurrentChatFromState();
  persistState();
  renderMessages();
  closeMessageActionPopup();
  await generateAssistantReply({ requestMessages, assistantMessage });
}

function setStatus(text) {
  els.statusText.textContent = text;
}

function showToast(message, type = "success") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  els.toastContainer.appendChild(toast);
  window.setTimeout(() => toast.remove(), 2600);
}

function persistState() {
  // Source of truth: local device file named `user`.
  persistUserFile();
}

window.addEventListener("beforeunload", () => {
  try {
    if (pendingUserFilePayload && hasFileSystemAccessApi()) {
      // best-effort flush request
      void flushPendingUserFileWrite();
    }
  } catch {
    // no-op
  }
});

function showCharacterEditor(character = null) {
  const normalized = character ? normalizeCharacterRecord(character) : {
    id: "",
    name: "",
    scenario: "",
    personality: "",
    greeting: "",
    alternateGreetings: [],
    avatarDataUrl: "",
  };
  els.characterEditorCard.dataset.editingId = character?.id || "";
  els.characterEditorCard.dataset.avatarDataUrl = normalized?.avatarDataUrl || "";
  els.characterNameInput.value = normalized?.name || "";
  els.characterScenarioInput.value = normalized?.scenario || "";
  els.characterPersonalityInput.value = normalized?.personality || "";
  els.characterGreetingInput.value = normalized?.greeting || "";
  if (els.characterAlternateGreetingsInput) {
    els.characterAlternateGreetingsInput.value = formatAlternateGreetingsForEditor(normalized?.alternateGreetings || []);
  }
  renderCharacterAvatarPreview(normalized?.avatarDataUrl || "");
  els.characterNameInput.focus();
}

function hideCharacterEditor() {
  delete els.characterEditorCard.dataset.editingId;
  delete els.characterEditorCard.dataset.avatarDataUrl;
  delete els.characterEditorCard.dataset.importedData;
  els.characterNameInput.value = "";
  els.characterScenarioInput.value = "";
  els.characterPersonalityInput.value = "";
  els.characterGreetingInput.value = "";
  if (els.characterAlternateGreetingsInput) {
    els.characterAlternateGreetingsInput.value = "";
  }
  els.characterAvatarInput.value = "";
  renderCharacterAvatarPreview("");
  switchView("charactersView");
}

async function onCharacterAvatarSelected(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const avatarDataUrl = await readFileAsDataUrl(file);
    els.characterEditorCard.dataset.avatarDataUrl = avatarDataUrl;
    renderCharacterAvatarPreview(avatarDataUrl);
  } catch {
    setStatus("Could not load avatar image");
    showToast("Avatar load failed", "error");
  }
}

function saveCharacterFromEditor() {
  const name = els.characterNameInput.value.trim();
  const scenario = els.characterScenarioInput.value.trim();
  const personality = els.characterPersonalityInput.value.trim();
  const greeting = els.characterGreetingInput.value.trim();
  const alternateGreetings = parseAlternateGreetings(els.characterAlternateGreetingsInput?.value || "", { fromEditor: true });
  const avatarDataUrl = els.characterEditorCard.dataset.avatarDataUrl || "";
  if (!name) {
    setStatus("Add a character name");
    showToast("Character save failed", "error");
    return;
  }

  const editingId = els.characterEditorCard.dataset.editingId;
  const character = editingId ? state.characters.find((item) => item.id === editingId) : null;
  const characterPayload = {
    name,
    scenario,
    personality,
    first_mes: greeting,
    alternate_greetings: alternateGreetings,
    description: personality || scenario || "",
    avatarDataUrl,
    spec: "chara_card_v2",
  };

  if (character) {
    character.name = name;
    character.scenario = scenario;
    character.personality = personality;
    character.greeting = greeting;
    character.alternateGreetings = alternateGreetings;
    character.description = personality || scenario || "";
    character.avatarDataUrl = avatarDataUrl;
    character.data = {
      ...(character.data && typeof character.data === "object" ? character.data : {}),
      ...characterPayload,
    };
    setStatus("Character updated");
    showToast("Character saved!", "success");
  } else {
    state.characters.unshift(normalizeCharacterRecord({
      id: crypto.randomUUID(),
      name,
      scenario,
      personality,
      greeting,
      alternateGreetings,
      description: personality || scenario || "",
      avatarDataUrl,
      data: characterPayload,
    }));
    setStatus("Character created");
    showToast("Character saved!", "success");
  }

  persistState();
  renderCharacters();
  renderImpersonationOptions();
  renderDrawerState();
  hideCharacterEditor();
  showToast("Character saved!", "success");
}

function renderCharacterAvatarPreview(avatarDataUrl) {
  els.characterAvatarPreview.classList.toggle("hidden", !avatarDataUrl);
  setElementBackgroundImage(els.characterAvatarPreview, avatarDataUrl);
}

function syncCustomizationInputs() {
  if (els.avatarSizeInput) els.avatarSizeInput.value = state.customization.avatarSize;
  if (els.defaultTextColorInput) els.defaultTextColorInput.value = state.customization.defaultTextColor;
  if (els.wrapRuleSearchInput) els.wrapRuleSearchInput.value = wrapRuleSearchQuery;
  if (els.googleFontSearchInput && state.customization.appFontSource === "google") {
    els.googleFontSearchInput.value = state.customization.appFontGoogleName || state.customization.appFontFamily || "";
  }
  if (els.uploadedFontMeta) {
    if (state.customization.appFontSource === "file" && state.customization.appFontFileName) {
      els.uploadedFontMeta.textContent = `Uploaded: ${state.customization.appFontFileName}`;
    } else if (state.customization.appFontSource === "google" && state.customization.appFontFamily) {
      els.uploadedFontMeta.textContent = `Google Font: ${state.customization.appFontFamily}`;
    } else {
      els.uploadedFontMeta.textContent = "Using system default font.";
    }
  }
  updateCurrentFontLabel();
}

function getCurrentFontDisplayName() {
  if (state.customization.appFontSource === "file") {
    return state.customization.appFontFamily || state.customization.appFontFileName || "Custom file font";
  }
  if (state.customization.appFontSource === "google") {
    return state.customization.appFontFamily || state.customization.appFontGoogleName || "Google Font";
  }
  return "System default (Inter stack)";
}

function updateCurrentFontLabel() {
  if (!els.currentFontName) return;
  els.currentFontName.textContent = getCurrentFontDisplayName();
}

function renderWrapRuleList() {
  if (!els.wrapRuleList) return;
  const query = wrapRuleSearchQuery.trim().toLowerCase();
  const rules = state.customization.wrapRules.filter((rule) => {
    if (!query) return true;
    const haystack = `${rule.title || ""} ${rule.start || ""} ${rule.end || ""}`.toLowerCase();
    return haystack.includes(query);
  });

  els.wrapRuleList.innerHTML = "";
  if (!rules.length) {
    const empty = document.createElement("p");
    empty.className = "entity-meta wrap-list-empty-state";
    empty.textContent = query ? "No wrap rules match your search." : "No wrap rules yet.";
    els.wrapRuleList.appendChild(empty);
    return;
  }

  rules.forEach((rule) => {
    const item = document.createElement("article");
    item.className = "entity-card";
    item.innerHTML = `
      <div class="entity-card-header">
        <h3>${escapeHtml(rule.title || `${rule.start}${rule.end}`)}</h3>
        <span class="entity-meta">${escapeHtml(rule.start)} ... ${escapeHtml(rule.end)}</span>
      </div>
      <div class="entity-card-actions">
        <span class="wrap-color-dot" aria-hidden="true" style="background: ${escapeAttribute(rule.color || "#c084fc")}"></span>
        <button class="secondary" type="button" data-action="edit-wrap">Edit</button>
      </div>
    `;
    item.querySelector('[data-action="edit-wrap"]')?.addEventListener("click", () => showWrapRuleEditor(rule.id));
    els.wrapRuleList.appendChild(item);
  });
}

function onWrapRuleSearchInput() {
  wrapRuleSearchQuery = els.wrapRuleSearchInput?.value || "";
  renderWrapRuleList();
}

function createWrapRule() {
  const rule = {
    id: crypto.randomUUID(),
    title: "New wrap rule",
    start: "(",
    end: ")",
    color: "#c084fc",
    enabled: true,
  };
  wrapRuleSearchQuery = "";
  if (els.wrapRuleSearchInput) els.wrapRuleSearchInput.value = "";
  state.customization.wrapRules.unshift(rule);
  persistState();
  renderWrapRuleList();
  showWrapRuleEditor(rule.id);
}

function showWrapRuleEditor(ruleId) {
  const rule = state.customization.wrapRules.find((item) => item.id === ruleId);
  if (!rule || !els.wrapRuleEditor) return;
  if (els.textWrappingSection?.classList.contains("hidden")) {
    els.textWrappingSection.classList.remove("hidden");
  }
  els.wrapRuleEditor.classList.remove("hidden");
  els.wrapRuleEditor.dataset.ruleId = rule.id;
  els.wrapRuleTitleInput.value = rule.title || "";
  els.wrapRuleStartInput.value = rule.start || "";
  els.wrapRuleEndInput.value = rule.end || "";
  els.wrapRuleColorInput.value = isValidHexColor(rule.color) ? rule.color : "#c084fc";
  els.wrapRuleHexInput.value = rule.color || "#c084fc";
  els.wrapRuleTitleInput?.focus();
}

function hideWrapRuleEditor() {
  if (!els.wrapRuleEditor) return;
  els.wrapRuleEditor.classList.add("hidden");
  delete els.wrapRuleEditor.dataset.ruleId;
  if (els.wrapRuleTitleInput) els.wrapRuleTitleInput.value = "";
  if (els.wrapRuleStartInput) els.wrapRuleStartInput.value = "";
  if (els.wrapRuleEndInput) els.wrapRuleEndInput.value = "";
  if (els.wrapRuleColorInput) els.wrapRuleColorInput.value = "#c084fc";
  if (els.wrapRuleHexInput) els.wrapRuleHexInput.value = "";
}

function saveWrapRuleFromEditor() {
  const rule = state.customization.wrapRules.find((item) => item.id === els.wrapRuleEditor.dataset.ruleId);
  if (!rule) return;
  const start = els.wrapRuleStartInput.value;
  const end = els.wrapRuleEndInput.value;
  if (!start || !end) {
    showToast("Wrap start and end are required", "error");
    return;
  }
  rule.start = start;
  rule.end = end;
  rule.color = isValidHexColor(els.wrapRuleHexInput.value.trim()) ? els.wrapRuleHexInput.value.trim() : els.wrapRuleColorInput.value;
  rule.title = els.wrapRuleTitleInput.value.trim() || `${rule.start}...${rule.end}`;
  persistState();
  renderWrapRuleList();
  applyCustomization();
  renderMessages();
  hideWrapRuleEditor();
  showToast("Wrap updated!", "success");
}

function deleteWrapRuleFromEditor() {
  const ruleId = els.wrapRuleEditor?.dataset.ruleId;
  if (!ruleId) return;
  state.customization.wrapRules = state.customization.wrapRules.filter((item) => item.id !== ruleId);
  persistState();
  renderWrapRuleList();
  applyCustomization();
  renderMessages();
  hideWrapRuleEditor();
  showToast("Wrap deleted", "success");
}

function toggleSection(element) {
  if (!element) return;
  element.classList.toggle("hidden");
  if (element.classList.contains("hidden")) return;
  if (element === els.chatCustomizationSection || element === els.textWrappingSection || element === els.textFontSection) {
    syncCustomizationInputs();
    renderWrapRuleList();
  }
  if (element === els.textFontSection) {
    toggleFontManager(true);
  }
}

function applyCustomization() {
  document.documentElement.style.setProperty("--chat-avatar-size", `${state.customization.avatarSize}px`);
  document.documentElement.style.setProperty("--chat-default-text-color", state.customization.defaultTextColor);
  applyWrapRuleCssVariables();
  applyAppFont();
  updateCurrentFontLabel();
}

function saveCustomization() {
  if (!els.avatarSizeInput || !els.defaultTextColorInput) {
    showToast("Customisation controls are unavailable", "error");
    return;
  }
  state.customization.avatarSize = Math.min(120, Math.max(24, Number.parseInt(els.avatarSizeInput.value, 10) || defaults.customization.avatarSize));
  state.customization.defaultTextColor = els.defaultTextColorInput.value || defaults.customization.defaultTextColor;
  persistState();
  syncCustomizationInputs();
  applyCustomization();
  renderMessages();
  showToast("Customisation saved!", "success");
  setStatus("Customisation saved");
}

function onDefaultTextColorLiveChange() {
  if (!els.defaultTextColorInput) return;
  state.customization.defaultTextColor = els.defaultTextColorInput.value || defaults.customization.defaultTextColor;
  applyCustomization();
  renderMessages();
  persistState();
}

function onWrapRuleColorInputChange() {
  if (!els.wrapRuleColorInput || !els.wrapRuleHexInput) return;
  const color = els.wrapRuleColorInput.value;
  els.wrapRuleHexInput.value = color;
  const ruleId = els.wrapRuleEditor?.dataset?.ruleId;
  if (!ruleId) return;
  const rule = state.customization.wrapRules.find((item) => item.id === ruleId);
  if (!rule) return;
  rule.color = color;
  applyCustomization();
  renderMessages();
  persistState();
}

function onWrapRuleHexInputChange() {
  if (!els.wrapRuleHexInput || !els.wrapRuleColorInput) return;
  const value = els.wrapRuleHexInput.value.trim();
  if (!isValidHexColor(value)) return;
  els.wrapRuleColorInput.value = value;
  const ruleId = els.wrapRuleEditor?.dataset?.ruleId;
  if (!ruleId) return;
  const rule = state.customization.wrapRules.find((item) => item.id === ruleId);
  if (!rule) return;
  rule.color = value;
  applyCustomization();
  renderMessages();
  persistState();
}

function applyWrapRuleCssVariables() {
  state.customization.wrapRules.forEach((rule, index) => {
    document.documentElement.style.setProperty(`--chat-wrap-color-${index}`, rule.color || "#c084fc");
  });
}

function openTextFontManager() {
  if (els.chatCustomizationSection?.classList.contains("hidden")) {
    els.chatCustomizationSection.classList.remove("hidden");
  }
  if (els.textFontSection?.classList.contains("hidden")) {
    els.textFontSection.classList.remove("hidden");
  }
  syncCustomizationInputs();
  toggleFontManager(true);
}

function toggleFontManager(show) {
  if (!els.fontManagerBackdrop) return;
  els.fontManagerBackdrop.classList.toggle("hidden", !show);
  if (show) {
    setInlineStatus(els.fontManagerStatus, "");
    syncCustomizationInputs();
  }
}

function isSupportedFontFile(fileName = "") {
  return /\.(ttf|otf|woff2?|TTF|OTF|WOFF2?)$/.test(String(fileName));
}

function getFontFormatFromFileName(fileName = "") {
  const lower = String(fileName).toLowerCase();
  if (lower.endsWith(".ttf")) return "truetype";
  if (lower.endsWith(".otf")) return "opentype";
  if (lower.endsWith(".woff")) return "woff";
  if (lower.endsWith(".woff2")) return "woff2";
  return "";
}

function extractFontFamilyFromFileName(fileName = "") {
  return String(fileName)
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[_-]+/g, " ")
    .trim();
}

function toSafeFontFamily(fontName = "") {
  const cleaned = String(fontName || "")
    .replace(/["'`]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || "Custom App Font";
}

function ensureAppFontStyleTag() {
  const existing = document.getElementById("appFontDynamicStyle");
  if (existing) {
    appFontStyleTag = existing;
    return existing;
  }
  const styleTag = document.createElement("style");
  styleTag.id = "appFontDynamicStyle";
  document.head.appendChild(styleTag);
  appFontStyleTag = styleTag;
  return styleTag;
}

function buildGoogleFontsCssUrl(fontName = "") {
  const token = String(fontName || "")
    .trim()
    .split(/\s+/)
    .map((part) => encodeURIComponent(part))
    .join("+");
  return `https://fonts.googleapis.com/css2?family=${token}&display=swap`;
}

function escapeCssValue(value) {
  return String(value || "")
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\n", " ");
}

function applyAppFont() {
  const fallbackFontStack = "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, \"Segoe UI\", sans-serif";
  const source = state.customization.appFontSource;
  let fontFamily = "";
  let fontCss = "";

  if (source === "file" && state.customization.appFontDataUrl) {
    fontFamily = toSafeFontFamily(state.customization.appFontFamily || extractFontFamilyFromFileName(state.customization.appFontFileName));
    const format = getFontFormatFromFileName(state.customization.appFontFileName);
    fontCss = `@font-face { font-family: "${escapeCssValue(fontFamily)}"; src: url("${escapeCssValue(state.customization.appFontDataUrl)}")${format ? ` format("${format}")` : ""}; font-display: swap; }`;
  } else if (source === "google" && (state.customization.appFontGoogleCss || state.customization.appFontGoogleName || state.customization.appFontFamily)) {
    fontFamily = toSafeFontFamily(state.customization.appFontFamily || state.customization.appFontGoogleName);
    if (state.customization.appFontGoogleCss) {
      fontCss = state.customization.appFontGoogleCss;
    } else if (state.customization.appFontGoogleName) {
      fontCss = `@import url("${escapeCssValue(buildGoogleFontsCssUrl(state.customization.appFontGoogleName))}");`;
    }
  }

  if (fontFamily && fontCss) {
    const styleTag = ensureAppFontStyleTag();
    styleTag.textContent = `${fontCss}\n:root { --app-font-family: "${escapeCssValue(fontFamily)}", ${fallbackFontStack}; }`;
    document.documentElement.style.setProperty("--app-font-family", `"${fontFamily}", ${fallbackFontStack}`);
  } else {
    document.documentElement.style.setProperty("--app-font-family", fallbackFontStack);
    if (appFontStyleTag) {
      appFontStyleTag.remove();
      appFontStyleTag = null;
    } else {
      document.getElementById("appFontDynamicStyle")?.remove();
    }
  }
}

async function onFontFileSelected(event) {
  const [file] = event?.target?.files || [];
  if (!file) return;
  if (!isSupportedFontFile(file.name)) {
    setInlineStatus(els.fontManagerStatus, "Unsupported file. Use .ttf, .otf, .woff, or .woff2.");
    showToast("Unsupported font file", "error");
    if (event?.target) event.target.value = "";
    return;
  }

  try {
    setInlineStatus(els.fontManagerStatus, "Loading font file...");
    const dataUrl = await readFileAsDataUrl(file);
    const family = toSafeFontFamily(extractFontFamilyFromFileName(file.name));

    state.customization.appFontSource = "file";
    state.customization.appFontFamily = family;
    state.customization.appFontFileName = file.name;
    state.customization.appFontDataUrl = dataUrl;
    state.customization.appFontGoogleName = "";
    state.customization.appFontGoogleCss = "";

    persistState();
    applyCustomization();
    syncCustomizationInputs();
    setInlineStatus(els.fontManagerStatus, `Applied \"${family}\" from file.`);
    showToast("Font applied", "success");
  } catch (error) {
    setInlineStatus(els.fontManagerStatus, error?.message || "Failed to read font file");
    showToast("Failed to load font file", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

async function addFontFromGoogleFonts() {
  const rawName = els.googleFontSearchInput?.value?.trim() || "";
  if (!rawName) {
    setInlineStatus(els.fontManagerStatus, "Enter a font name to search Google Fonts.");
    showToast("Enter a Google Font name", "error");
    return;
  }

  try {
    setInlineStatus(els.fontManagerStatus, `Checking Google Fonts for \"${rawName}\"...`);
    const url = buildGoogleFontsCssUrl(rawName);
    const response = await fetch(url, { method: "GET" });
    const css = await response.text();

    if (!response.ok || !css.includes("@font-face")) {
      throw new Error("No matching Google Font found.");
    }

    const familyMatch = css.match(/font-family:\s*['\"]([^'\"]+)['\"]/i);
    const resolvedFamily = toSafeFontFamily(familyMatch?.[1] || rawName);

    state.customization.appFontSource = "google";
    state.customization.appFontFamily = resolvedFamily;
    state.customization.appFontGoogleName = rawName;
    state.customization.appFontGoogleCss = css;
    state.customization.appFontFileName = "";
    state.customization.appFontDataUrl = "";

    persistState();
    applyCustomization();
    syncCustomizationInputs();
    setInlineStatus(els.fontManagerStatus, `Added Google Font: ${resolvedFamily}`);
    showToast("Google Font added", "success");
  } catch (error) {
    setInlineStatus(els.fontManagerStatus, error?.message || "Could not load that Google Font.");
    showToast("Google Font not found", "error");
  }
}

function resetAppFont() {
  state.customization.appFontSource = defaults.customization.appFontSource;
  state.customization.appFontFamily = defaults.customization.appFontFamily;
  state.customization.appFontFileName = "";
  state.customization.appFontDataUrl = "";
  state.customization.appFontGoogleName = "";
  state.customization.appFontGoogleCss = "";
  persistState();
  applyCustomization();
  syncCustomizationInputs();
  setInlineStatus(els.fontManagerStatus, "Font reset to system default.");
  showToast("Font reset", "success");
}

function renderImpersonationOptions() {
  const current = els.impersonateSelect.value;
  els.impersonateSelect.innerHTML = '<option value="">Speak as yourself</option>';
  state.characters.forEach((character) => {
    const option = document.createElement("option");
    option.value = character.id;
    option.textContent = character.name;
    option.selected = character.id === current;
    els.impersonateSelect.appendChild(option);
  });
  updateImpersonationButton();
}

function toggleImpersonationPicker() {
  const shouldShow = els.impersonateSelect.classList.contains("hidden");
  els.impersonateSelect.classList.toggle("hidden", !shouldShow);
  els.impersonateToggleBtn.classList.toggle("active", shouldShow);
  if (shouldShow) {
    els.impersonateSelect.focus();
  }
}

function onImpersonationChange() {
  updateImpersonationButton();
  els.impersonateSelect.classList.add("hidden");
  els.impersonateToggleBtn.classList.remove("active");
}

function updateImpersonationButton() {
  const selectedId = els.impersonateSelect.value;
  const selectedCharacter = state.characters.find((character) => character.id === selectedId);
  els.impersonateToggleBtn.title = selectedCharacter ? `Speaking as ${selectedCharacter.name}` : "Speak as yourself";
}

function toggleChatDrawer(show) {
  toggleSettings(false);
  if (show && state.activeView !== "chatsView") {
    return;
  }
  els.chatDrawer.classList.toggle("hidden", !show);
  if (!show) {
    els.presetPicker.classList.add("hidden");
    els.providerPicker.classList.add("hidden");
    hideDrawerPresetEditor();
    hideDrawerProviderEditor();
  }
}

function onChatTitleInput() {
  state.currentChatTitle = els.chatTitleInput.value.trim() || "New chat";
  const chat = getCurrentChat();
  if (chat) {
    chat.title = state.currentChatTitle;
    syncCurrentChatFromState();
  }
  persistState();
  renderRecentChats();
}

function getActivePreset() {
  return state.presets.find((preset) => preset.id === state.activePresetId) || state.presets[0] || null;
}

function renderDrawerState() {
  els.chatTitleInput.value = state.currentChatTitle || "New chat";
  const preset = getActivePreset();
  const provider = state.providers.find((item) => item.id === state.activeProviderId);
  els.currentPresetLabel.textContent = `Current preset: ${preset?.name || "Default preset"}`;
  els.currentProviderLabel.textContent = `Current provider: ${provider?.name || "None selected"}`;
  renderDrawerPresets();
  renderDrawerProviders();
  renderDrawerLorebooks();
  updateImpersonationButton();
}

function getCurrentChatCharacter() {
  const chat = getCurrentChat();
  if (!chat?.characterId) return null;
  return state.characters.find((character) => character.id === chat.characterId) || null;
}

function readAvatarCandidate(value) {
  if (!value) return "";
  if (typeof value === "string") {
    const normalized = value.trim();
    if (!normalized || ["none", "null", "undefined"].includes(normalized.toLowerCase())) return "";
    return normalized;
  }
  if (typeof value === "object") {
    const nested = value.dataUrl
      || value.dataURL
      || value.url
      || value.src
      || value.path
      || value.avatar;
    return readAvatarCandidate(nested);
  }
  return "";
}

function getCharacterAvatarUrl(character = null) {
  if (!character || typeof character !== "object") return "";
  return readAvatarCandidate(character.avatarDataUrl)
    || readAvatarCandidate(character.avatar)
    || readAvatarCandidate(character.avatarUrl)
    || readAvatarCandidate(character.avatar_url)
    || readAvatarCandidate(character.data?.avatarDataUrl)
    || readAvatarCandidate(character.data?.avatar)
    || readAvatarCandidate(character.data?.avatarUrl)
    || readAvatarCandidate(character.data?.avatar_url)
    || "";
}

function setElementBackgroundImage(element, rawUrl) {
  if (!element) return;
  const url = String(rawUrl || "").trim();
  if (!url) {
    element.style.backgroundImage = "";
    return;
  }
  element.style.backgroundImage = `url(${JSON.stringify(url)})`;
}

function getMessageAvatar(msg) {
  if (msg.role === "assistant") {
    return getCharacterAvatarUrl(getCurrentChatCharacter()) || readAvatarCandidate(msg?.avatarDataUrl);
  }
  if (msg.role === "user") {
    const selectedCharacter = state.characters.find((character) => character.id === els.impersonateSelect.value);
    return getCharacterAvatarUrl(selectedCharacter);
  }
  return "";
}

function getMessageDisplayRole(msg) {
  if (msg.role === "assistant") {
    return getCurrentChatCharacter()?.name || "assistant";
  }
  if (msg.role === "user") {
    const selectedCharacter = state.characters.find((character) => character.id === els.impersonateSelect.value);
    return selectedCharacter?.name || "user";
  }
  return msg.role;
}

function formatMessageContent(content) {
  let html = escapeHtml(content || "");
  state.customization.wrapRules.forEach((rule, index) => {
    if (rule?.enabled === false) return;

    const startRaw = String(rule.start || "");
    const endRaw = String(rule.end || "");
    if (!startRaw || !endRaw) return;

    const escapedStart = escapeHtml(startRaw);
    const escapedEnd = escapeHtml(endRaw);
    const start = escapeRegex(escapedStart);
    const end = escapeRegex(escapedEnd);
    if (!start || !end) return;

    const className = `format-wrap-${index}`;
    if (startRaw === "*" && endRaw === "*") {
      html = html.replace(new RegExp(`${start}([^\n]+?)${end}`, "g"), `<span class="${className}" style="color: var(--chat-wrap-color-${index})">$1</span>`);
      return;
    }

    html = html.replace(new RegExp(`${start}([^]+?)${end}`, "g"), `<span class="${className}" style="color: var(--chat-wrap-color-${index})">${escapedStart}$1${escapedEnd}</span>`);
  });
  return html;
}

function resolveCurrentUserName() {
  const selectedCharacter = state.characters.find((character) => character.id === els.impersonateSelect?.value);
  return selectedCharacter?.name?.trim() || "User";
}

function getMacroLocalVarsStore() {
  const chatId = state.currentChatId || "";
  if (!chatId) return {};
  if (!state.macroChatVars[chatId] || typeof state.macroChatVars[chatId] !== "object") {
    state.macroChatVars[chatId] = {};
  }
  return state.macroChatVars[chatId];
}

function normalizeMacroValue(value) {
  if (value === undefined || value === null) return "";
  return String(value);
}

function parseDiceFormula(formula) {
  const raw = String(formula || "").trim();
  if (!raw) return null;
  const normalized = /^\d+$/.test(raw) ? `1d${raw}` : raw;
  const match = normalized.match(/^(\d*)d(\d+)([+-]\d+)?$/i);
  if (!match) return null;
  const count = Number(match[1] || 1);
  const sides = Number(match[2]);
  const mod = Number(match[3] || 0);
  if (!Number.isFinite(count) || !Number.isFinite(sides) || !Number.isFinite(mod)) return null;
  if (count < 1 || sides < 1 || count > 1000 || sides > 100000) return null;
  return { count, sides, mod };
}

function rollDice(formula) {
  const parsed = parseDiceFormula(formula);
  if (!parsed) return "";
  let total = parsed.mod;
  for (let i = 0; i < parsed.count; i += 1) {
    total += 1 + Math.floor(Math.random() * parsed.sides);
  }
  return String(total);
}

function simpleHash(value) {
  let hash = 0;
  const str = String(value || "");
  for (let i = 0; i < str.length; i += 1) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash);
}

function splitMacroArgs(raw = "") {
  const input = String(raw || "");
  const parts = [];
  let current = "";
  for (let i = 0; i < input.length; i += 1) {
    const char = input[i];
    const next = input[i + 1];
    if (char === "\\" && next) {
      current += next;
      i += 1;
      continue;
    }
    if (char === ":" && next === ":") {
      parts.push(current.trim());
      current = "";
      i += 1;
      continue;
    }
    current += char;
  }
  parts.push(current.trim());
  return parts;
}

function splitMacroListArg(raw) {
  const value = String(raw || "");
  if (value.includes("::")) {
    return splitMacroArgs(value).map((item) => item.trim()).filter(Boolean);
  }
  const parts = [];
  let current = "";
  for (let i = 0; i < value.length; i += 1) {
    const char = value[i];
    const next = value[i + 1];
    if (char === "\\" && next) {
      current += next;
      i += 1;
      continue;
    }
    if (char === ",") {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  parts.push(current.trim());
  return parts.filter(Boolean);
}


function getLastMessageByRole(role = "") {
  const list = Array.isArray(state.messages) ? state.messages : [];
  for (let i = list.length - 1; i >= 0; i -= 1) {
    const msg = list[i];
    if (!msg || (msg.role !== "user" && msg.role !== "assistant")) continue;
    if (!role || msg.role === role) return msg;
  }
  return null;
}

function evaluateMacroCondition(rawCondition, context) {
  const resolved = applyMacrosToText(String(rawCondition || ""), context.character, context.userName, {
    vars: context.vars,
    skipScopedIf: true,
  }).trim();
  if (!resolved) return false;
  const lowered = resolved.toLowerCase();
  if (["0", "false", "off", "no", "null", "undefined"].includes(lowered)) return false;
  return true;
}

function applyMacrosToText(text, character = null, userNameOverride = "", options = {}) {
  const sourceText = String(text || "");
  const charRef = character || getCurrentChatCharacter();
  const charName = charRef?.name?.trim() || "Character";
  const userName = userNameOverride || resolveCurrentUserName();
  const lastMessage = getLastMessageByRole("");
  const lastUserMessage = getLastMessageByRole("user");
  const lastCharMessage = getLastMessageByRole("assistant");
  const preset = getActivePreset();
  const instruct = preset?.instruct || {};
  const systemPrompt = preset?.systemPrompt?.trim() || state.systemPrompt || "";
  const localVars = getMacroLocalVarsStore();
  const globalVars = state.macroGlobalVars || {};
  const creatorNotes = charRef?.data?.creator_notes || charRef?.data?.creatorNotes || "";
  const charVersion = charRef?.data?.character_version || charRef?.data?.char_version || charRef?.data?.version || "";
  const mesExamplesRaw = charRef?.data?.mes_example || charRef?.data?.mesExamples || "";
  const hasMessages = Array.isArray(state.messages) && state.messages.length > 0;
  const latestSwipeMessage = hasMessages
    ? [...state.messages].reverse().find((msg) => msg?.role === "assistant" && Array.isArray(msg.swipes) && msg.swipes.length)
    : null;
  const currentSwipeId = latestSwipeMessage
    ? clampSwipeIndex(Number(latestSwipeMessage.swipe_id ?? 0), latestSwipeMessage.swipes.length)
    : 0;
  const lastSwipeId = latestSwipeMessage ? Math.max(0, latestSwipeMessage.swipes.length - 1) : 0;
  const isMobile = typeof window !== "undefined" && typeof window.matchMedia === "function"
    ? window.matchMedia("(max-width: 768px)").matches
    : false;

  const baseVars = {
    char: charName,
    user: userName,
    charIfNotGroup: charName,
    group: charName,
    groupNotMuted: charName,
    notChar: userName,
    persona: "",
    charPrompt: "",
    charInstruction: "",
    charDescription: charRef?.data?.description || charRef?.personality || "",
    description: charRef?.data?.description || charRef?.personality || "",
    charPersonality: charRef?.personality || charRef?.data?.personality || "",
    personality: charRef?.personality || charRef?.data?.personality || "",
    charScenario: charRef?.scenario || charRef?.data?.scenario || "",
    scenario: charRef?.scenario || charRef?.data?.scenario || "",
    charFirstMessage: charRef?.greeting || charRef?.data?.first_mes || "",
    greeting: charRef?.greeting || charRef?.data?.first_mes || "",
    charCreatorNotes: creatorNotes,
    creatorNotes,
    charVersion,
    version: charVersion,
    char_version: charVersion,
    mesExamplesRaw,
    mesExamples: mesExamplesRaw,
    charDepthPrompt: "",
    original: charRef?.data?.original || "",
    model: state.selectedModel || "",
    systemPrompt,
    defaultSystemPrompt: systemPrompt,
    instructSystem: systemPrompt,
    instructSystemPrompt: systemPrompt,
    chatStart: preset?.chatStart || "",
    exampleSeparator: preset?.exampleSeparator || "",
    chatSeparator: preset?.exampleSeparator || "",
    maxPrompt: String(preset?.maxTokens ?? ""),
    maxPromptTokens: String(preset?.maxTokens ?? ""),
    maxContext: String(preset?.maxTokens ?? ""),
    maxContextTokens: String(preset?.maxTokens ?? ""),
    maxResponse: String(preset?.maxTokens ?? ""),
    maxResponseTokens: String(preset?.maxTokens ?? ""),
    isMobile: isMobile ? "true" : "false",
    lastGenerationType: String(state.lastGenerationType || ""),
    lastMessage: lastMessage?.content || "",
    lastUserMessage: lastUserMessage?.content || "",
    lastCharMessage: lastCharMessage?.content || "",
    lastMessageId: lastMessage ? String(Math.max(0, state.messages.length - 1)) : "",
    firstIncludedMessageId: hasMessages ? "0" : "",
    firstDisplayedMessageId: hasMessages ? "0" : "",
    currentSwipeId: String(currentSwipeId),
    lastSwipeId: String(lastSwipeId),
    allChatRange: state.messages?.length ? `0-${Math.max(0, state.messages.length - 1)}` : "",
    instructStoryStringPrefix: instruct.storyStringPrefix || "",
    instructStoryStringSuffix: instruct.storyStringSuffix || "",
    instructUserPrefix: instruct.inputSequence || "",
    instructInput: instruct.inputSequence || "",
    instructUserSuffix: instruct.inputSuffix || "",
    instructAssistantPrefix: instruct.outputSequence || "",
    instructOutput: instruct.outputSequence || "",
    instructAssistantSuffix: instruct.outputSuffix || instruct.stopSequence || "",
    instructSeparator: instruct.outputSuffix || instruct.stopSequence || "",
    instructSystemPrefix: instruct.systemSequence || "",
    instructSystemSuffix: instruct.systemSuffix || "",
    instructFirstAssistantPrefix: instruct.firstOutputSequence || instruct.outputSequence || "",
    instructFirstOutputPrefix: instruct.firstOutputSequence || instruct.outputSequence || "",
    instructLastAssistantPrefix: instruct.lastOutputSequence || instruct.outputSequence || "",
    instructLastOutputPrefix: instruct.lastOutputSequence || instruct.outputSequence || "",
    instructStop: instruct.stopSequence || "",
    instructUserFiller: "",
    instructSystemInstructionPrefix: instruct.systemSequence || "",
    instructFirstUserPrefix: instruct.firstInputSequence || instruct.inputSequence || "",
    instructFirstInput: instruct.firstInputSequence || instruct.inputSequence || "",
    instructLastUserPrefix: instruct.lastInputSequence || instruct.inputSequence || "",
    instructLastInput: instruct.lastInputSequence || instruct.inputSequence || "",
  };

  const vars = {
    ...baseVars,
    ...(options.vars && typeof options.vars === "object" ? options.vars : {}),
  };

  let output = sourceText;

  const resolveNamedMacro = (nameRaw, argRaw, fullToken = "") => {
    const macroName = String(nameRaw || "").trim();
    const macro = macroName.toLowerCase();
    const arg = String(argRaw || "");

    if (macro === "space") {
      const [count] = splitMacroArgs(arg);
      return " ".repeat(Math.max(1, Number(count || 1)));
    }
    if (macro === "newline") {
      const [count] = splitMacroArgs(arg);
      return "\n".repeat(Math.max(1, Number(count || 1)));
    }
    if (["noop", "trim", "banned", "outlet"].includes(macro)) return "";
    if (macro === "//" || macro === "comment") return "";
    if (macro === "datetimeformat") {
      const now = new Date();
      const format = arg.trim();
      if (!format) return now.toLocaleString();
      return format
        .replace(/YYYY/g, String(now.getFullYear()))
        .replace(/MM/g, String(now.getMonth() + 1).padStart(2, "0"))
        .replace(/DD/g, String(now.getDate()).padStart(2, "0"))
        .replace(/HH/g, String(now.getHours()).padStart(2, "0"))
        .replace(/mm/g, String(now.getMinutes()).padStart(2, "0"))
        .replace(/ss/g, String(now.getSeconds()).padStart(2, "0"));
    }
    if (macro === "time") return new Date().toLocaleTimeString();
    if (macro === "date") return new Date().toLocaleDateString();
    if (macro === "weekday") return new Date().toLocaleDateString(undefined, { weekday: "long" });
    if (macro === "isotime") return new Date().toISOString().slice(11, 16);
    if (macro === "isodate") return new Date().toISOString().slice(0, 10);
    if (macro === "idleduration" || macro === "idle_duration") return "just now";
    if (macro === "timediff") {
      const [left = "", right = ""] = splitMacroArgs(arg);
      const a = new Date(left.trim());
      const b = new Date(right.trim());
      if (Number.isNaN(a.getTime()) || Number.isNaN(b.getTime())) return "";
      const ms = Math.abs(a.getTime() - b.getTime());
      const mins = Math.floor(ms / 60000);
      if (mins < 60) return `${mins} minute${mins === 1 ? "" : "s"}`;
      const hours = Math.floor(mins / 60);
      if (hours < 24) return `${hours} hour${hours === 1 ? "" : "s"}`;
      const days = Math.floor(hours / 24);
      return `${days} day${days === 1 ? "" : "s"}`;
    }

    if ([
      "setvar", "addvar", "incvar", "decvar", "getvar", "hasvar", "varexists", "deletevar", "flushvar",
      "setglobalvar", "addglobalvar", "incglobalvar", "decglobalvar", "getglobalvar", "hasglobalvar", "globalvarexists", "deleteglobalvar", "flushglobalvar",
    ].includes(macro)) {
      const parts = [macroName, ...splitMacroArgs(arg)];
      const normalized = parts[0].toLowerCase();
      const name = String(parts[1] || "").trim();
      const rawValue = parts.slice(2).join("::").trim();
      const target = normalized.includes("global") ? globalVars : localVars;
      if (!name && !["incvar", "decvar", "incglobalvar", "decglobalvar"].includes(normalized)) return "";

      const current = target[name];
      const currentNum = Number(current);
      const valueNum = Number(rawValue);
      const canAddNum = Number.isFinite(currentNum) && Number.isFinite(valueNum);

      if (normalized === "setvar" || normalized === "setglobalvar") {
        target[name] = rawValue;
        return "";
      }
      if (normalized === "addvar" || normalized === "addglobalvar") {
        target[name] = canAddNum ? String(currentNum + valueNum) : `${normalizeMacroValue(current)}${rawValue}`;
        return "";
      }
      if (normalized === "incvar" || normalized === "incglobalvar") {
        target[name] = String((Number.isFinite(currentNum) ? currentNum : 0) + 1);
        return target[name];
      }
      if (normalized === "decvar" || normalized === "decglobalvar") {
        target[name] = String((Number.isFinite(currentNum) ? currentNum : 0) - 1);
        return target[name];
      }
      if (normalized === "getvar" || normalized === "getglobalvar") return normalizeMacroValue(target[name]);
      if (["hasvar", "varexists", "hasglobalvar", "globalvarexists"].includes(normalized)) {
        return Object.prototype.hasOwnProperty.call(target, name) ? "true" : "false";
      }
      if (["deletevar", "flushvar", "deleteglobalvar", "flushglobalvar"].includes(normalized)) {
        delete target[name];
        return "";
      }
      return "";
    }

    if (macro === "roll") return rollDice(arg);
    if (macro === "reverse") return Array.from(arg).reverse().join("");
    if (macro === "input") return String(els.messageInput?.value || "");
    if (macro === "hasextension") {
      const [extensionName = ""] = splitMacroArgs(arg);
      return extensionName ? "false" : "false";
    }
    if (macro === "random" || macro === "pick") {
      const items = splitMacroListArg(arg);
      if (!items.length) return "";
      if (macro === "pick") {
        const seed = simpleHash(`${state.currentChatId || ""}::${fullToken || `${macroName}::${arg}`}`);
        return items[seed % items.length] || "";
      }
      return items[Math.floor(Math.random() * items.length)] || "";
    }

    if (macro === "if") {
      const [cond = "", yesVal = "", noVal = ""] = splitMacroArgs(arg);
      return evaluateMacroCondition(cond, { character: charRef, userName, vars }) ? String(yesVal) : String(noVal);
    }

    if (Object.prototype.hasOwnProperty.call(vars, macroName)) {
      return String(vars[macroName] ?? "");
    }

    const key = Object.keys(vars).find((item) => item.toLowerCase() === macro);
    if (key) return String(vars[key] ?? "");

    return "";
  };

  if (!options.skipScopedIf) {
    output = output.replace(/{{\s*if\s+([^}]+)}}([\s\S]*?){{\s*\/if\s*}}/gi, (_, conditionRaw, body) => {
      const [thenBranch, elseBranch = ""] = String(body || "").split(/{{\s*else\s*}}/i);
      return evaluateMacroCondition(conditionRaw, { character: charRef, userName, vars }) ? thenBranch : elseBranch;
    });
  }

  let changedInLastPass = false;
  for (let pass = 0; pass < 16; pass += 1) {
    const before = output;
    output = output.replace(/{{\s*([^{}]+?)\s*}}/g, (full, innerRaw) => {
      const inner = String(innerRaw || "").trim();
      if (!inner) return "";
      if (/^#if\s+/i.test(inner) || /^\/if$/i.test(inner) || /^else$/i.test(inner)) return "";
      const [name, ...rest] = splitMacroArgs(inner);
      const arg = rest.join("::");
      return resolveNamedMacro(name, arg, full);
    });

    output = output
      .replace(/<USER>/gi, userName)
      .replace(/<BOT>/gi, charName)
      .replace(/<CHAR>/gi, charName);

    changedInLastPass = before !== output;
    if (!changedInLastPass) break;
  }

  if (changedInLastPass && /{{\s*[^{}]+\s*}}/.test(output)) {
    console.warn("MaiTavern macro engine: reached max-pass limit while resolving macros", {
      sample: output.slice(0, 240),
    });
  }

  persistState();
  return output;
}

function parseAlternateGreetings(value, options = {}) {
  const fromEditor = Boolean(options?.fromEditor);

  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }

  const text = String(value || "").replace(/\r\n/g, "\n").trim();
  if (!text) return [];

  if (/\n\s*---\s*\n/.test(text)) {
    return text
      .split(/\n\s*---\s*\n/g)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  if (!fromEditor && /^\s*\[[\s\S]*\]\s*$/.test(text)) {
    try {
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        return parsed.map((item) => String(item || "").trim()).filter(Boolean);
      }
    } catch {
      // ignore and continue with legacy parsing
    }
  }

  if (fromEditor) {
    return [text];
  }

  const lines = text.split("\n");
  const hasBlankLines = lines.some((line) => !line.trim());
  if (hasBlankLines) {
    return [text];
  }

  return lines.map((item) => item.trim()).filter(Boolean);
}

function formatAlternateGreetingsForEditor(value) {
  const greetings = parseAlternateGreetings(value);
  if (!greetings.length) return "";
  return greetings.join("\n\n---\n\n");
}

function getCharacterGreetingVariants(character) {
  const primary = String(character?.greeting || character?.data?.first_mes || "").trim();
  const alternates = parseAlternateGreetings(character?.alternateGreetings || character?.data?.alternate_greetings || []);
  return [primary, ...alternates].filter(Boolean);
}

function normalizeCharacterRecord(item = {}) {
  const data = item?.data && typeof item.data === "object" ? item.data : item;
  const name = data?.name || item?.name || "";
  const scenario = data?.scenario || item?.scenario || "";
  const personality = data?.personality || item?.personality || data?.description || item?.description || "";
  const greeting = data?.first_mes || item?.first_mes || data?.greeting || item?.greeting || "";
  const alternateGreetings = parseAlternateGreetings(
    item?.alternateGreetings
    || data?.alternate_greetings
    || data?.alternateGreetings
    || item?.alternate_greetings
    || ""
  );
  const avatarDataUrl = readAvatarCandidate(
    item?.avatarDataUrl
    || item?.avatar
    || item?.avatarUrl
    || item?.avatar_url
    || data?.avatarDataUrl
    || data?.avatar
    || data?.avatarUrl
    || data?.avatar_url
    || ""
  );
  const fallbackDescription = name || "Imported character";
  return {
    id: item?.id || crypto.randomUUID(),
    name,
    scenario,
    personality,
    greeting,
    alternateGreetings,
    description: personality || scenario || item?.description || fallbackDescription,
    avatarDataUrl,
    data: {
      ...(data && typeof data === "object" ? data : {}),
      name,
      scenario,
      personality,
      first_mes: greeting,
      alternate_greetings: alternateGreetings,
      description: personality || scenario || item?.description || "",
      avatarDataUrl,
      avatar: avatarDataUrl,
    },
  };
}

function normalizeChatRecord(chat = {}) {
  const lorebookSettingsSource = chat?.lorebookSettings && typeof chat.lorebookSettings === "object"
    ? chat.lorebookSettings
    : defaults.lorebookSettings;

  return {
    ...chat,
    lorebookIds: Array.isArray(chat.lorebookIds)
      ? chat.lorebookIds.map((item) => String(item || "").trim()).filter(Boolean)
      : [],
    lorebookSettings: {
      scanDepth: Math.min(64, Math.max(1, Number.parseInt(lorebookSettingsSource?.scanDepth, 10) || defaults.lorebookSettings.scanDepth)),
      tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(lorebookSettingsSource?.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget)),
      recursiveScanning: lorebookSettingsSource?.recursiveScanning !== false,
    },
    messages: Array.isArray(chat.messages) ? chat.messages.map((message) => normalizeMessageRecord(message)) : [],
  };
}

function normalizeMessageRecord(message = {}) {
  const normalized = { ...message };
  if (Array.isArray(normalized.swipes)) {
    normalized.swipes = normalized.swipes.map((item) => String(item ?? ""));
    normalized.swipe_id = clampSwipeIndex(Number(normalized.swipe_id ?? 0), normalized.swipes.length);
    normalized.content = normalized.swipes[normalized.swipe_id] || normalized.content || "";
  }
  return normalized;
}

function clampSwipeIndex(index, length) {
  if (!Number.isFinite(index)) return 0;
  if (!length || length < 1) return 0;
  return Math.min(Math.max(0, index), length - 1);
}

function getMessageSwipeContent(message) {
  if (!Array.isArray(message?.swipes) || !message.swipes.length) {
    return message?.content || "";
  }
  const index = clampSwipeIndex(Number(message.swipe_id ?? 0), message.swipes.length);
  return message.swipes[index] || message.content || "";
}

function setMessageSwipeIndex(messageIndex, nextIndex) {
  const message = state.messages[messageIndex];
  if (!message || !Array.isArray(message.swipes) || message.swipes.length < 2) return;
  const clamped = clampSwipeIndex(nextIndex, message.swipes.length);
  message.swipe_id = clamped;
  message.content = message.swipes[clamped] || message.content || "";
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
}

function createUserFilePayload() {
  return {
    app: "maitavern",
    schemaVersion: USER_FILE_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    state: JSON.parse(JSON.stringify(state)),
  };
}

function hasFileSystemAccessApi() {
  return typeof fetch === "function";
}

async function writeUserFileToDisk(payload) {
  try {
    const response = await fetch(USER_FILE_ENDPOINT, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(text || `HTTP ${response.status}`);
    }

    deviceStorageUnavailable = false;
  } catch (error) {
    deviceStorageUnavailable = true;
    throw new Error(`Local device storage unavailable. Start MaiTavern server (${error?.message || "request failed"})`);
  }
}

async function readUserFileFromDisk() {
  try {
    const response = await fetch(USER_FILE_ENDPOINT, {
      method: "GET",
      headers: { "Accept": "application/json" },
    });

    if (response.status === 404) {
      throw new Error("User file not found");
    }

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(text || `HTTP ${response.status}`);
    }

    const parsed = await response.json();
    if (!parsed || typeof parsed !== "object") {
      throw new Error("User file is invalid");
    }

    deviceStorageUnavailable = false;
    return { parsed };
  } catch (error) {
    deviceStorageUnavailable = true;
    throw new Error(`Local device storage unavailable. Start MaiTavern server (${error?.message || "request failed"})`);
  }
}

async function flushPendingUserFileWrite() {
  if (!pendingUserFilePayload || userFileWriteInFlight) return;
  if (!hasFileSystemAccessApi()) return;

  userFileWriteInFlight = true;
  try {
    const payload = pendingUserFilePayload;
    pendingUserFilePayload = null;
    await writeUserFileToDisk(payload);
    setStatus(`Saved to local file: ${USER_FILE_NAME}`);
  } catch {
    // Keep payload for retry on next save.
  } finally {
    userFileWriteInFlight = false;
    if (pendingUserFilePayload) {
      if (userFileWriteTimer) clearTimeout(userFileWriteTimer);
      userFileWriteTimer = setTimeout(() => {
        void flushPendingUserFileWrite();
      }, 300);
    }
  }
}

function persistUserFile() {
  const payload = createUserFilePayload();
  if (!hasFileSystemAccessApi()) return;

  pendingUserFilePayload = payload;
  if (userFileWriteTimer) {
    clearTimeout(userFileWriteTimer);
  }
  userFileWriteTimer = setTimeout(() => {
    void flushPendingUserFileWrite();
  }, 220);
}

function loadUserFile() {
  return null;
}

function loadLegacyStateFromStorage() {
  return structuredClone(defaults);
}

async function exportUserDataZip() {
  try {
    const payloadText = JSON.stringify(createUserFilePayload(), null, 2);
    const zipBytes = createSingleFileZip("maitavern-user.json", payloadText);
    downloadBlob("maitavern-data.zip", new Blob([zipBytes], { type: "application/zip" }));
    setDataStatus("Export completed");
    setStatus("Data exported");
    showToast("Data exported", "success");
  } catch (error) {
    setDataStatus(error?.message || "Export failed");
    setStatus("Data export failed");
    showToast("Data export failed", "error");
  }
}

async function importUserDataZip(event) {
  const [file] = event?.target?.files || [];
  if (!file) return;

  try {
    setDataStatus("Importing...");
    const buffer = await file.arrayBuffer();
    const zipEntry = extractSingleFileFromZip(buffer);
    const parsed = JSON.parse(zipEntry.text);
    const importedState = parsed?.state && typeof parsed.state === "object" ? parsed.state : parsed;

    state = {
      ...structuredClone(defaults),
      ...importedState,
    };

    ensureCollections();
    syncInputsFromState();
    renderModelOptions();
    applyCustomization();
    renderMessages();
    renderCharacters();
    renderProviders();
    renderPresets();
    renderLorebooks();
    renderRecentChats();
    renderImpersonationOptions();
    renderDrawerState();
    syncCustomizationInputs();
    syncLorebookSettingsInputs();
    renderWrapRuleList();
    applyWelcomeState();
    switchView(state.activeView || "homeView");
    updateChatUiVisibility();
    persistState();

    setDataStatus(`Imported ${zipEntry.name}`);
    setStatus("Data imported");
    showToast("Data imported", "success");
  } catch (error) {
    setDataStatus(error?.message || "Import failed");
    setStatus("Data import failed");
    showToast("Data import failed", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function createSingleFileZip(filename, textContent) {
  const encoder = new TextEncoder();
  const nameBytes = encoder.encode(filename);
  const dataBytes = encoder.encode(textContent);
  const crc = crc32(dataBytes);

  const localHeader = new Uint8Array(30 + nameBytes.length);
  const localView = new DataView(localHeader.buffer);
  localView.setUint32(0, 0x04034b50, true);
  localView.setUint16(4, 20, true);
  localView.setUint16(6, 0, true);
  localView.setUint16(8, 0, true);
  localView.setUint16(10, 0, true);
  localView.setUint16(12, 0, true);
  localView.setUint32(14, crc, true);
  localView.setUint32(18, dataBytes.length, true);
  localView.setUint32(22, dataBytes.length, true);
  localView.setUint16(26, nameBytes.length, true);
  localView.setUint16(28, 0, true);
  localHeader.set(nameBytes, 30);

  const centralHeader = new Uint8Array(46 + nameBytes.length);
  const centralView = new DataView(centralHeader.buffer);
  centralView.setUint32(0, 0x02014b50, true);
  centralView.setUint16(4, 20, true);
  centralView.setUint16(6, 20, true);
  centralView.setUint16(8, 0, true);
  centralView.setUint16(10, 0, true);
  centralView.setUint16(12, 0, true);
  centralView.setUint16(14, 0, true);
  centralView.setUint32(16, crc, true);
  centralView.setUint32(20, dataBytes.length, true);
  centralView.setUint32(24, dataBytes.length, true);
  centralView.setUint16(28, nameBytes.length, true);
  centralView.setUint16(30, 0, true);
  centralView.setUint16(32, 0, true);
  centralView.setUint16(34, 0, true);
  centralView.setUint16(36, 0, true);
  centralView.setUint32(38, 0, true);
  centralView.setUint32(42, 0, true);
  centralHeader.set(nameBytes, 46);

  const endHeader = new Uint8Array(22);
  const endView = new DataView(endHeader.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, 1, true);
  endView.setUint16(10, 1, true);
  endView.setUint32(12, centralHeader.length, true);
  endView.setUint32(16, localHeader.length + dataBytes.length, true);
  endView.setUint16(20, 0, true);

  return concatUint8Arrays([localHeader, dataBytes, centralHeader, endHeader]);
}

function extractSingleFileFromZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);

  const signatureOffset = findZipLocalHeaderOffset(bytes);
  if (signatureOffset < 0) {
    throw new Error("Invalid ZIP file");
  }

  const signature = view.getUint32(signatureOffset, true);
  if (signature !== 0x04034b50) {
    throw new Error("ZIP local file header not found");
  }

  const compressionMethod = view.getUint16(signatureOffset + 8, true);
  if (compressionMethod !== 0) {
    throw new Error("Unsupported ZIP compression method");
  }

  const compressedSize = view.getUint32(signatureOffset + 18, true);
  const fileNameLength = view.getUint16(signatureOffset + 26, true);
  const extraLength = view.getUint16(signatureOffset + 28, true);

  const nameStart = signatureOffset + 30;
  const dataStart = nameStart + fileNameLength + extraLength;
  const dataEnd = dataStart + compressedSize;

  if (dataEnd > bytes.length) {
    throw new Error("Corrupted ZIP file data");
  }

  const decoder = new TextDecoder();
  const name = decoder.decode(bytes.slice(nameStart, nameStart + fileNameLength));
  const text = decoder.decode(bytes.slice(dataStart, dataEnd));
  return { name, text };
}

function findZipLocalHeaderOffset(bytes) {
  for (let i = 0; i < bytes.length - 3; i += 1) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
      return i;
    }
  }
  return -1;
}

function concatUint8Arrays(arrays) {
  const total = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  arrays.forEach((arr) => {
    out.set(arr, offset);
    offset += arr.length;
  });
  return out;
}

let crcTable = null;

function getCrcTable() {
  if (crcTable) return crcTable;
  crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    crcTable[i] = c >>> 0;
  }
  return crcTable;
}

function crc32(bytes) {
  const table = getCrcTable();
  let crc = 0 ^ (-1);
  for (let i = 0; i < bytes.length; i += 1) {
    crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xff];
  }
  return (crc ^ (-1)) >>> 0;
}

async function loadState() {
  // Prefer local device file first.
  if (hasFileSystemAccessApi()) {
    try {
      await pickAndLoadUserFile();
      return state;
    } catch {
      // User may cancel picker; fall back safely.
    }
  }

  return loadLegacyStateFromStorage();
}

function migrateLegacyChats() {
  if (state.chats.length) {
    if (state.currentChatId && !state.chats.some((chat) => chat.id === state.currentChatId)) {
      state.currentChatId = state.chats[0]?.id || "";
    }
    if (!state.currentChatId && state.chats[0]) {
      state.currentChatId = state.chats[0].id;
      state.currentChatTitle = state.chats[0].title || "New chat";
      state.messages = Array.isArray(state.chats[0].messages) ? state.chats[0].messages.map((message) => ({ ...message })) : [];
    }
    return;
  }

  const usefulMessages = Array.isArray(state.messages)
    ? state.messages.filter((message) => message.role === "user" || message.role === "assistant")
    : [];

  if (!usefulMessages.length && !state.recentChats.length) {
    state.messages = [];
    return;
  }

  const migratedChat = {
    id: crypto.randomUUID(),
    title: state.currentChatTitle || state.recentChats[0]?.title || "New chat",
    messages: usefulMessages.map((message) => ({ ...message })),
    preview: (state.recentChats[0]?.preview || usefulMessages[usefulMessages.length - 1]?.content || "").slice(0, 80),
    updatedAt: state.recentChats[0]?.updatedAt || new Date().toISOString(),
  };

  state.chats = [migratedChat];
  state.recentChats = [{
    id: migratedChat.id,
    title: migratedChat.title,
    preview: migratedChat.preview,
    updatedAt: migratedChat.updatedAt,
  }];
  state.currentChatId = migratedChat.id;
  state.currentChatTitle = migratedChat.title;
  state.messages = migratedChat.messages.map((message) => ({ ...message }));
}

function autoResizeTextarea(textarea) {
  textarea.style.height = "auto";
  textarea.style.height = Math.min(textarea.scrollHeight, 180) + "px";
}

function downloadJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  downloadBlob(filename, blob);
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function slugify(value) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "item";
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}

function decodeBase64Unicode(base64Value) {
  const normalized = String(base64Value || "").trim().replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new TextDecoder().decode(bytes);
}

function normalizePngChunkKeyword(rawKeyword = "") {
  return String(rawKeyword || "").replace(/\0+$/g, "").trim().toLowerCase();
}

function parseEmbeddedCharacterJson(rawText = "") {
  const text = String(rawText || "").trim();
  if (!text) return null;

  const directCandidates = [text];
  if (text.startsWith("\ufeff")) {
    directCandidates.push(text.replace(/^\ufeff/, ""));
  }

  for (const candidate of directCandidates) {
    try {
      const parsed = JSON.parse(candidate);
      return parsed?.data || parsed;
    } catch {
      // continue
    }
  }

  try {
    const decoded = decodeBase64Unicode(text);
    const parsed = JSON.parse(decoded);
    return parsed?.data || parsed;
  } catch {
    // continue
  }

  return null;
}

function normalizeImportedCharacter(item, fileName = "") {
  const normalized = normalizeCharacterRecord(item || {});
  if (!normalized.name) {
    normalized.name = fileName.replace(/\.json$/i, "") || "Imported character";
    normalized.data.name = normalized.name;
  }
  if (!normalized.description) {
    normalized.description = normalized.name || "Imported character";
  }
  return {
    ...normalized,
    id: crypto.randomUUID(),
  };
}

async function importCharacterCardFromImage(file) {
  const avatarDataUrl = await readFileAsDataUrl(file);
  const embedded = await extractCharacterCardFromPng(file).catch(() => null);
  if (embedded) {
    const normalized = normalizeImportedCharacter(embedded, file.name);
    normalized.avatarDataUrl = avatarDataUrl;
    if (normalized.data && typeof normalized.data === "object") {
      normalized.data.avatarDataUrl = avatarDataUrl;
    }
    return normalized;
  }
  const name = file.name.replace(/\.[^.]+$/, "") || "Imported character";
  return {
    id: crypto.randomUUID(),
    name,
    description: "Imported from avatar image",
    avatarDataUrl,
    data: {
      name,
      description: "Imported from avatar image",
      avatarDataUrl,
      spec: "chara_card_v2",
    },
  };
}

async function extractCharacterCardFromPng(file) {
  const buffer = await readFileAsArrayBuffer(file);
  const bytes = new Uint8Array(buffer);

  if (bytes.length < 8) {
    throw new Error("PNG file is too small");
  }

  const pngSignature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  const isPng = pngSignature.every((value, index) => bytes[index] === value);
  if (!isPng) {
    throw new Error("File is not a valid PNG");
  }

  const decoder = new TextDecoder();
  const view = new DataView(buffer);
  let offset = 8;
  const textCandidates = [];

  while (offset + 8 <= bytes.length) {
    const length = view.getUint32(offset, false);
    const typeStart = offset + 4;
    const typeEnd = typeStart + 4;
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    const crcEnd = dataEnd + 4;

    if (crcEnd > bytes.length) break;

    const chunkType = decoder.decode(bytes.slice(typeStart, typeEnd));
    const chunkData = bytes.slice(dataStart, dataEnd);

    if (chunkType === "tEXt") {
      const separatorIndex = chunkData.indexOf(0);
      if (separatorIndex > 0) {
        const keyword = normalizePngChunkKeyword(decoder.decode(chunkData.slice(0, separatorIndex)));
        const textValue = decoder.decode(chunkData.slice(separatorIndex + 1));
        textCandidates.push({ keyword, textValue });
      }
    } else if (chunkType === "iTXt") {
      const null1 = chunkData.indexOf(0);
      if (null1 > 0) {
        const keyword = normalizePngChunkKeyword(decoder.decode(chunkData.slice(0, null1)));
        const compressionFlag = chunkData[null1 + 1] ?? 0;
        const compressionMethod = chunkData[null1 + 2] ?? 0;
        let cursor = null1 + 3;
        const null2 = chunkData.indexOf(0, cursor);
        if (null2 >= 0) {
          cursor = null2 + 1;
          const null3 = chunkData.indexOf(0, cursor);
          if (null3 >= 0) {
            cursor = null3 + 1;
            if (compressionFlag === 0 && compressionMethod === 0) {
              const textValue = decoder.decode(chunkData.slice(cursor));
              textCandidates.push({ keyword, textValue });
            }
          }
        }
      }
    }

    if (chunkType === "IEND") break;
    offset = crcEnd;
  }

  const prioritized = [
    ...textCandidates.filter((item) => item.keyword === "chara"),
    ...textCandidates.filter((item) => item.keyword !== "chara"),
  ];

  for (const candidate of prioritized) {
    const parsed = parseEmbeddedCharacterJson(candidate.textValue);
    if (parsed && typeof parsed === "object") {
      return parsed;
    }
  }

  throw new Error("No embedded character data found in image");
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;
  return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
}

function isValidHexColor(value) {
  return /^#(?:[0-9a-fA-F]{3}){1,2}$/.test(String(value || "").trim());
}

function escapeAttribute(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("'", "&#39;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
