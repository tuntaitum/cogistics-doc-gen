// services.js — shared by index.html (lists services) and service.html (lists one
// service's tools). Both are built from services.json: to add a service or a tool,
// edit that file; no HTML changes needed.
//
// Names and descriptions are set with textContent. The SVG icons below are fixed
// strings of ours (never data from services.json), so innerHTML is safe for them.

const ICONS = {
  truck: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h11v10H3z"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="7" cy="17.5" r="1.8"/><circle cx="17" cy="17.5" r="1.8"/></svg>',
  leaf: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M5 19c0-8 5-13 14-14 0 9-5 14-13 14"/><path d="M5 19c3-4 6-6 9-8"/></svg>',
  sheet: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 10h16M4 15h16M10 4v16"/></svg>',
  letter: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/></svg>',
  report: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M7 3h7l5 5v13H7z"/><path d="M14 3v5h5M10 13h6M10 17h6"/></svg>',
};
const DEFAULT_ICON = "report";
const ARROW = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"/></svg>';

async function loadServices() {
  const res = await fetch("services.json", { cache: "no-cache" });
  if (!res.ok) throw new Error(res.status);
  return (await res.json()).services;
}

// entries: [{name, description, href, icon}] -> one navy button each
function renderCards(container, entries) {
  for (const e of entries) {
    const a = document.createElement("a");
    a.className = "service-card";
    a.href = e.href;

    const icon = document.createElement("span");
    icon.className = "service-card__icon";
    icon.setAttribute("aria-hidden", "true");
    icon.innerHTML = ICONS[e.icon] || ICONS[DEFAULT_ICON];

    const text = document.createElement("span");
    text.className = "service-card__text";
    const name = document.createElement("h2");
    name.className = "service-card__name";
    name.textContent = e.name;
    const desc = document.createElement("p");
    desc.className = "service-card__desc";
    desc.textContent = e.description;
    text.append(name, desc);

    const arrow = document.createElement("span");
    arrow.className = "service-card__arrow";
    arrow.setAttribute("aria-hidden", "true");
    arrow.innerHTML = ARROW;

    a.append(icon, text, arrow);
    container.appendChild(a);
  }
}
