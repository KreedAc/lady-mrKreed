/* Piano alimentare — app statica, nessuna dipendenza.
   I dati arrivano da data/plan.json (o gia' incorporati in window.__PLAN_DATA__). */

(function () {
  "use strict";

  var app = document.getElementById("app");
  var backBtn = document.getElementById("back");
  var brandText = document.getElementById("brand-text");
  var DATA = null;

  var EMOJI = {
    colazione: "🍳",
    pranzo: "🥗",
    merenda: "🍎",
    cena: "🍝",
    "dopo cena": "🌙",
    spuntino: "🥛",
  };

  var CAT_EMOJI = {
    "frutta e verdura": "🥬",
    "carne e salumi": "🥩",
    pesce: "🐟",
    "latticini e uova": "🧀",
    "pane, cereali e dispensa": "🍞",
    "dispensa e condimenti": "🫙",
  };

  // ------------------------------------------------------------ helpers

  function h(tag, attrs, children) {
    var el = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (key) {
      var value = attrs[key];
      if (value === null || value === undefined || value === false) return;
      if (key === "class") el.className = value;
      else if (key === "text") el.textContent = value;
      else if (key === "html") el.innerHTML = value;
      else if (key.indexOf("on") === 0) el.addEventListener(key.slice(2), value);
      else el.setAttribute(key, value);
    });
    (children || []).forEach(function (child) {
      if (child === null || child === undefined || child === false) return;
      el.appendChild(typeof child === "string" ? document.createTextNode(child) : child);
    });
    return el;
  }

  function kcal(value) {
    if (value === null || value === undefined || value === "") return "";
    return Math.round(value) + " kcal";
  }

  function grams(value) {
    if (value === null || value === undefined) return "";
    return value + " g";
  }

  function chevron() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("class", "chevron");
    svg.setAttribute("aria-hidden", "true");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M9 5l7 7-7 7");
    svg.appendChild(path);
    return svg;
  }

  function searchIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    var circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("cx", "11");
    circle.setAttribute("cy", "11");
    circle.setAttribute("r", "7");
    var line = document.createElementNS("http://www.w3.org/2000/svg", "path");
    line.setAttribute("d", "M20 20l-4-4");
    svg.appendChild(circle);
    svg.appendChild(line);
    return svg;
  }

  function badge(text, accent) {
    return text ? h("span", { class: "badge" + (accent ? " accent" : ""), text: text }) : null;
  }

  function personByKey(key) {
    return DATA.people.filter(function (p) {
      return p.key === key;
    })[0];
  }

  function recipeBySlug(person, slug) {
    return person.recipes.filter(function (r) {
      return r.slug === slug;
    })[0];
  }

  function normalize(text) {
    return (text || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  /* Le kcal delle ricette a volte sono un numero, a volte una stima tipo "≈430 kcal". */
  function kcalLabel(recipe) {
    if (recipe.kcal_label && /[^\d.,\s]/.test(recipe.kcal_label)) return recipe.kcal_label;
    return kcal(recipe.kcal);
  }

  /* Gli ingredienti nel foglio sono una riga sola separata da "·": li divido in elenco. */
  function splitIngredients(text) {
    if (!text) return [];
    return text
      .split("·")
      .map(function (part) {
        return part.trim().replace(/\.$/, "");
      })
      .filter(Boolean);
  }

  function setChrome(title, accent, showBack) {
    brandText.textContent = title;
    document.body.setAttribute("data-accent", accent || "teal");
    backBtn.hidden = !showBack;
    document.title = title + " — Piano alimentare";
  }

  // ------------------------------------------------------------ viste

  function viewHome() {
    setChrome("Piano alimentare", "teal", false);

    var hero = h("div", { class: "hero" }, [
      h("h1", { text: "Piano alimentare" }),
      h("p", { text: "Menu della settimana, ricette e lista della spesa — sempre aggiornati dal foglio Excel." }),
    ]);

    var cards = DATA.people.map(function (person) {
      var stats = person.stats;
      var plan = person.plan;
      var settimane = plan.weeks.length;
      return h(
        "a",
        { class: "person-card", href: "#/p/" + person.key, "data-accent": person.accent },
        [
          h("div", { class: "avatar", text: person.name.charAt(0) }),
          h("h2", { text: person.name }),
          h("p", {
            class: "lede",
            text:
              (plan.kcal_target ? "~" + Math.round(plan.kcal_target) + " kcal al giorno · " : "") +
              (plan.meals_per_day ? plan.meals_per_day + " pasti" : ""),
          }),
          h("div", { class: "figures" }, [
            badge(stats.days + (stats.days === 1 ? " giorno" : " giorni"), true),
            badge(settimane > 1 ? settimane + " settimane" : "1 settimana"),
            badge(stats.recipes + " ricette"),
          ]),
        ]
      );
    });

    var totaleRicette = DATA.people.reduce(function (acc, p) {
      return acc + p.recipes.length;
    }, 0);

    var quick = [
      { href: "#/ricettario", emoji: "📖", title: "Ricettario", sub: totaleRicette + " ricette di Giovanni e Rosalia" },
      { href: "#/spesa", emoji: "🛒", title: "Lista della spesa", sub: "Con le spunte che restano salvate" },
      { href: "#/prova", emoji: "🧪", title: "Ricette da provare", sub: DATA.to_try.recipes.length + " idee non ancora nel piano" },
    ].map(function (item) {
      return h("a", { class: "quick", href: item.href }, [
        h("span", { class: "emoji", text: item.emoji }),
        h("div", {}, [h("strong", { text: item.title }), h("span", { text: item.sub })]),
      ]);
    });

    return h("div", {}, [hero, h("div", { class: "people" }, cards), h("div", { class: "quick-links" }, quick)]);
  }

  function viewPerson(key, weekIndex, dayIndex) {
    var person = personByKey(key);
    if (!person) return viewNotFound();

    setChrome(person.name, person.accent, true);

    var weeks = person.plan.weeks;
    var week = weeks[weekIndex] || weeks[0];
    var day = week.days[dayIndex] || week.days[0];

    var nodes = [];

    nodes.push(
      h("div", { class: "hero", style: "padding-bottom:16px" }, [
        h("h1", { text: person.name }),
        h("p", { text: person.plan.subtitle || "" }),
      ])
    );

    if (weeks.length > 1) {
      nodes.push(
        h(
          "div",
          { class: "tabs", role: "tablist" },
          weeks.map(function (w, i) {
            return h("button", {
              class: "tab",
              role: "tab",
              type: "button",
              "aria-selected": String(i === weekIndex),
              text: w.name,
              onclick: function () {
                go("#/p/" + key + "/" + i + "/0");
              },
            });
          })
        )
      );
    }

    nodes.push(
      h(
        "div",
        { class: "days" },
        week.days.map(function (d, i) {
          return h(
            "button",
            {
              class: "day-chip",
              type: "button",
              "aria-current": String(i === dayIndex),
              onclick: function () {
                go("#/p/" + key + "/" + weekIndex + "/" + i);
              },
            },
            [h("b", { text: d.name.slice(0, 3) }), h("small", { text: kcal(d.total) })]
          );
        })
      )
    );

    nodes.push(
      h("div", { class: "day-head" }, [
        h("h2", { text: day.name }),
        h("div", { class: "kcal-total" }, [
          document.createTextNode(kcal(day.total) + " "),
          h("small", { text: "totali" }),
        ]),
      ])
    );

    nodes.push(
      h(
        "div",
        { class: "stack" },
        day.meals.map(function (meal) {
          return mealCard(person, meal);
        })
      )
    );

    nodes.push(
      h("div", { class: "quick-links" }, [
        h("a", { class: "quick", href: "#/ricettario/" + key }, [
          h("span", { class: "emoji", text: "📖" }),
          h("div", {}, [
            h("strong", { text: "Tutte le ricette" }),
            h("span", { text: person.recipes.length + " piatti di " + person.name }),
          ]),
        ]),
        h("a", { class: "quick", href: "#/spesa" }, [
          h("span", { class: "emoji", text: "🛒" }),
          h("div", {}, [h("strong", { text: "Lista della spesa" }), h("span", { text: "Quantità per il periodo" })]),
        ]),
      ])
    );

    (person.plan.notes || []).forEach(function (note) {
      nodes.push(
        h("div", { class: "note", style: "margin-top:14px" }, [
          note.title ? h("h4", { text: note.title }) : null,
          h("p", { text: note.body || "" }),
        ])
      );
    });

    (person.extras || []).forEach(function (extra) {
      nodes.push(
        h("div", { class: "note", style: "margin-top:14px" }, [
          h("h4", { text: extra.title }),
          h(
            "ul",
            { class: "ing-list", style: "margin-top:8px" },
            splitIngredients(extra.body).map(function (line) {
              return h("li", { text: line });
            })
          ),
        ])
      );
    });

    return h("div", {}, nodes);
  }

  function mealCard(person, meal) {
    var recipe = meal.recipe;
    var slug = recipe && recipe.slug;
    var body = [];

    if (recipe) {
      var inner = [h("div", { class: "dish-title", text: recipe.title })];
      if (slug) {
        inner.push(chevron());
        body.push(h("a", { class: "dish", href: "#/r/" + person.key + "/" + slug }, inner));
      } else {
        inner.push(h("span", { class: "no-recipe", text: "ricetta non disponibile" }));
        body.push(h("div", { class: "dish" }, inner));
      }
    }

    if (meal.items && meal.items.length) {
      body.push(
        h(
          "table",
          { class: "items" },
          [
            h(
              "tbody",
              {},
              meal.items.map(function (item) {
                return h("tr", {}, [
                  h("td", { text: item.food }),
                  h("td", { class: "qty", text: grams(item.qty) }),
                  h("td", { class: "kcal", text: kcal(item.kcal) }),
                ]);
              })
            ),
          ]
        )
      );
    }

    var label = normalize(meal.name);
    return h("section", { class: "meal" }, [
      h("div", { class: "meal-head" }, [
        h("span", { class: "emoji", text: EMOJI[label] || "🍽️" }),
        h("h3", { text: meal.name }),
        badge(kcal(meal.kcal), true),
      ]),
      h("div", { class: "meal-body" }, body),
    ]);
  }

  function viewRecipe(key, slug) {
    var person = personByKey(key);
    if (!person) return viewNotFound();
    var recipe = recipeBySlug(person, slug);
    if (!recipe) return viewNotFound();

    setChrome(recipe.name, person.accent, true);

    var nodes = [];
    nodes.push(
      h("div", { class: "recipe-head" }, [
        h("h1", { text: recipe.name }),
        h("div", { class: "figures" }, [
          badge(kcalLabel(recipe), true),
          badge(recipe.when),
          badge(recipe.category),
          badge(person.name),
        ]),
      ])
    );

    var ingredienti = splitIngredients(recipe.ingredients);
    if (ingredienti.length) {
      nodes.push(
        h("section", { class: "panel" }, [
          h("h3", { text: "Ingredienti · 1 porzione" }),
          h(
            "ul",
            { class: "ing-list" },
            ingredienti.map(function (line) {
              return h("li", { text: line });
            })
          ),
        ])
      );
    }

    if (recipe.plan_items && recipe.plan_items.length) {
      nodes.push(
        h("section", { class: "panel" }, [
          h("h3", { text: "Ingredienti del pasto" }),
          h(
            "table",
            { class: "items" },
            [
              h(
                "tbody",
                {},
                recipe.plan_items.map(function (item) {
                  return h("tr", {}, [
                    h("td", { text: item.food }),
                    h("td", { class: "qty", text: grams(item.qty) }),
                    h("td", { class: "kcal", text: kcal(item.kcal) }),
                  ]);
                })
              ),
            ]
          ),
        ])
      );
    }

    if (recipe.prep) {
      nodes.push(
        h("section", { class: "panel" }, [h("h3", { text: "Preparazione" }), h("p", { text: recipe.prep })])
      );
    }

    if (recipe.used_in && recipe.used_in.length) {
      nodes.push(
        h("section", { class: "panel" }, [
          h("h3", { text: "Nel piano" }),
          h(
            "div",
            { class: "used-in" },
            recipe.used_in.map(function (use) {
              var etichetta = (person.plan.weeks.length > 1 ? use.week + " · " : "") + use.day + " · " + use.meal;
              return h("span", { class: "badge", text: etichetta });
            })
          ),
        ])
      );
    }

    nodes.push(
      h("div", { class: "quick-links" }, [
        h("a", { class: "quick", href: "#/ricettario/" + key }, [
          h("span", { class: "emoji", text: "📖" }),
          h("div", {}, [h("strong", { text: "Tutte le ricette" }), h("span", { text: "Ricettario di " + person.name })]),
        ]),
        h("a", { class: "quick", href: "#/p/" + key }, [
          h("span", { class: "emoji", text: "📅" }),
          h("div", {}, [h("strong", { text: "Torna al piano" }), h("span", { text: "Menu di " + person.name })]),
        ]),
      ])
    );

    var wrapper = h("div", {}, nodes);
    wrapper.className = "stack";
    return wrapper;
  }

  /* Ricettario: senza persona mostra tutto, con la persona solo le sue ricette. */
  function viewRicettario(key) {
    var person = key ? personByKey(key) : null;
    if (key && !person) return viewNotFound();
    var persone = person ? [person] : DATA.people;
    var totale = persone.reduce(function (acc, p) {
      return acc + p.recipes.length;
    }, 0);

    setChrome(person ? "Ricette · " + person.name : "Ricettario", person ? person.accent : "teal", true);

    var lista = h("div", {});
    var input = h("input", {
      type: "search",
      placeholder: "Cerca un piatto o un ingrediente…",
      "aria-label": "Cerca ricetta",
    });

    function render(query) {
      lista.innerHTML = "";
      var q = normalize(query);
      var trovate = 0;

      persone.forEach(function (p) {
        var ricette = p.recipes.filter(function (r) {
          if (!q) return true;
          return normalize(r.name + " " + r.ingredients + " " + r.prep).indexOf(q) !== -1;
        });
        if (!ricette.length) return;
        trovate += ricette.length;

        var categorie = [];
        ricette.forEach(function (r) {
          if (categorie.indexOf(r.category) === -1) categorie.push(r.category);
        });

        categorie.forEach(function (categoria) {
          var titolo = persone.length > 1 ? p.name + " · " + categoria : categoria;
          lista.appendChild(h("h2", { class: "section-title", text: titolo }));
          lista.appendChild(
            h(
              "div",
              { class: "card-list", "data-accent": p.accent },
              ricette
                .filter(function (r) {
                  return r.category === categoria;
                })
                .map(function (r) {
                  return h("a", { class: "recipe-card", href: "#/r/" + p.key + "/" + r.slug }, [
                    h("div", { class: "info" }, [
                      h("strong", { text: r.name }),
                      h("span", { text: r.ingredients || r.when || r.prep }),
                    ]),
                    badge(kcalLabel(r), true),
                    chevron(),
                  ]);
                })
            )
          );
        });
      });

      if (!trovate) lista.appendChild(h("p", { class: "empty", text: "Nessuna ricetta trovata." }));
    }

    input.addEventListener("input", function () {
      render(input.value);
    });
    render("");

    var filtri = [{ key: "", name: "Tutte" }].concat(
      DATA.people.map(function (p) {
        return { key: p.key, name: p.name };
      })
    );

    return h("div", {}, [
      h("div", { class: "hero", style: "padding-bottom:14px" }, [
        h("h1", { text: "Ricettario" }),
        h("p", { text: totale + " ricette" + (person ? " di " + person.name : " di Giovanni e Rosalia") }),
      ]),
      h(
        "div",
        { class: "tabs", role: "tablist" },
        filtri.map(function (f) {
          return h("button", {
            class: "tab",
            role: "tab",
            type: "button",
            "aria-selected": String(f.key === (key || "")),
            text: f.name,
            onclick: function () {
              go("#/ricettario" + (f.key ? "/" + f.key : ""));
            },
          });
        })
      ),
      h("label", { class: "search" }, [searchIcon(), input]),
      lista,
    ]);
  }

  function viewToTry() {
    setChrome("Da provare", "teal", true);
    var ricette = DATA.to_try.recipes;

    var categorie = [];
    ricette.forEach(function (r) {
      if (categorie.indexOf(r.category) === -1) categorie.push(r.category);
    });

    var nodes = [
      h("div", { class: "hero", style: "padding-bottom:8px" }, [
        h("h1", { text: "Da provare" }),
        h("p", { text: DATA.to_try.subtitle || "" }),
      ]),
    ];

    categorie.forEach(function (categoria) {
      nodes.push(h("h2", { class: "section-title", text: categoria }));
      nodes.push(
        h(
          "div",
          { class: "card-list" },
          ricette
            .filter(function (r) {
              return r.category === categoria;
            })
            .map(function (r) {
              return h("article", { class: "panel" }, [
                h("div", { class: "figures", style: "display:flex;gap:8px;align-items:baseline" }, [
                  h("h3", { style: "font-size:18px;text-transform:none;letter-spacing:0;color:var(--text);margin:0 auto 0 0;font-family:var(--serif)", text: r.name }),
                  badge(kcalLabel(r), true),
                ]),
                h("p", { style: "margin-top:10px", text: r.ingredients }),
                r.prep ? h("p", { style: "margin-top:10px;color:var(--muted)", text: r.prep }) : null,
              ]);
            })
        )
      );
    });

    return h("div", {}, nodes);
  }

  function viewSpesa(groupIndex, separate) {
    setChrome("Lista della spesa", "teal", true);
    var unite = DATA.shopping.merged || [];
    // di default le liste unite (Giovanni + Rosalia settimana per settimana)
    var separato = separate || !unite.length;
    var gruppi = separato ? DATA.shopping.groups : unite;
    if (!gruppi.length) return h("p", { class: "empty", text: "Nessuna lista della spesa nel file." });

    var indice = Math.min(groupIndex || 0, gruppi.length - 1);
    var gruppo = gruppi[indice];
    document.body.setAttribute("data-accent", gruppo.person === "rosalia" ? "rose" : "teal");

    var chiave = "spesa:" + (gruppo.slug || indice);
    var spuntati = {};
    try {
      spuntati = JSON.parse(localStorage.getItem(chiave) || "{}");
    } catch (err) {
      spuntati = {};
    }

    var contatore = h("span", {});
    var totale = gruppo.categories.reduce(function (acc, c) {
      return acc + c.items.length;
    }, 0);

    function aggiornaContatore() {
      var fatti = Object.keys(spuntati).filter(function (k) {
        return spuntati[k];
      }).length;
      contatore.textContent = fatti + " di " + totale + " nel carrello";
    }

    function salva() {
      try {
        localStorage.setItem(chiave, JSON.stringify(spuntati));
      } catch (err) {
        /* modalita' privata: pazienza, le spunte non si conservano */
      }
      aggiornaContatore();
    }

    var nodes = [
      h("div", { class: "hero", style: "padding-bottom:12px" }, [
        h("h1", { text: "Lista della spesa" }),
        h("p", {
          text: separato
            ? "Liste separate, una per persona"
            : "Giovanni e Rosalia insieme, prodotti uguali sommati",
        }),
      ]),
    ];

    if (gruppi.length > 1) {
      nodes.push(
        h(
          "div",
          { class: "tabs", role: "tablist" },
          gruppi.map(function (g, i) {
            var nome = separato
              ? (g.person ? g.person.charAt(0).toUpperCase() + g.person.slice(1) : g.name) + " · " + g.label
              : g.name;
            return h("button", {
              class: "tab",
              role: "tab",
              type: "button",
              "aria-selected": String(i === indice),
              text: nome,
              onclick: function () {
                go("#/spesa/" + (separato ? "sep/" : "") + i);
              },
            });
          })
        )
      );
    }

    if (gruppo.sources && gruppo.sources.length) {
      nodes.push(
        h(
          "p",
          { class: "sources" },
          [document.createTextNode("Unisce: " + gruppo.sources.join(" + "))]
        )
      );
    }

    var reset = h("button", {
      class: "link-btn",
      type: "button",
      text: "Svuota le spunte",
      onclick: function () {
        spuntati = {};
        salva();
        render();
      },
    });

    var cambia = unite.length
      ? h("button", {
          class: "link-btn",
          type: "button",
          text: separato ? "Unisci le liste" : "Vedi separate",
          onclick: function () {
            go("#/spesa/" + (separato ? "" : "sep/") + "0");
          },
        })
      : null;

    nodes.push(h("div", { class: "shop-toolbar" }, [contatore, cambia, reset]));

    var contenitore = h("div", {});
    nodes.push(contenitore);

    function render() {
      contenitore.innerHTML = "";
      gruppo.categories.forEach(function (categoria) {
        var emoji = CAT_EMOJI[normalize(categoria.name)] || "•";
        contenitore.appendChild(h("h2", { class: "section-title", text: emoji + "  " + categoria.name }));
        contenitore.appendChild(
          h(
            "div",
            { class: "shop-items" },
            categoria.items.map(function (item) {
              var box = h("input", { type: "checkbox", checked: spuntati[item.id] ? "checked" : null });
              var dettaglio = (item.parts || [])
                .map(function (p) {
                  return p.person + " " + p.qty;
                })
                .join(" · ");
              var riga = h("label", { class: "shop-item" + (spuntati[item.id] ? " done" : "") }, [
                box,
                h("span", { class: "name" }, [
                  document.createTextNode(item.name),
                  dettaglio ? h("small", { text: dettaglio }) : null,
                ]),
                h("span", { class: "qty", text: item.qty }),
              ]);
              box.addEventListener("change", function () {
                spuntati[item.id] = box.checked;
                riga.classList.toggle("done", box.checked);
                salva();
              });
              return riga;
            })
          )
        );
      });
      aggiornaContatore();
    }

    render();
    return h("div", {}, nodes);
  }

  function viewNotFound() {
    setChrome("Non trovato", "teal", true);
    return h("div", { class: "empty" }, [
      h("p", { text: "Questa pagina non esiste." }),
      h("p", {}, [h("a", { href: "#/", text: "Torna alla home", style: "color:var(--accent-text)" })]),
    ]);
  }

  // ------------------------------------------------------------ router

  function go(hash) {
    if (location.hash === hash) render();
    else location.hash = hash;
  }

  function render() {
    var parts = (location.hash || "#/").replace(/^#\/?/, "").split("/").filter(Boolean);
    var vista;

    switch (parts[0]) {
      case undefined:
        vista = viewHome();
        break;
      case "p":
        vista = viewPerson(parts[1], parseInt(parts[2], 10) || 0, parseInt(parts[3], 10) || 0);
        break;
      case "r":
        vista = viewRecipe(parts[1], parts[2]);
        break;
      case "ricettario":
        vista = viewRicettario(parts[1]);
        break;
      case "spesa":
        vista =
          parts[1] === "sep"
            ? viewSpesa(parseInt(parts[2], 10) || 0, true)
            : viewSpesa(parseInt(parts[1], 10) || 0, false);
        break;
      case "prova":
        vista = viewToTry();
        break;
      default:
        vista = viewNotFound();
    }

    app.innerHTML = "";
    app.appendChild(vista);
    window.scrollTo(0, 0);

    // la settimana e il giorno selezionati devono restare visibili anche se fuori scroll
    Array.prototype.forEach.call(
      app.querySelectorAll('[aria-current="true"], [aria-selected="true"]'),
      function (attivo) {
        if (attivo.scrollIntoView) attivo.scrollIntoView({ block: "nearest", inline: "center" });
      }
    );
  }

  backBtn.addEventListener("click", function () {
    if (history.length > 1) history.back();
    else go("#/");
  });

  // ------------------------------------------------------------ tema

  (function tema() {
    var salvato = null;
    try {
      salvato = localStorage.getItem("tema");
    } catch (err) {
      /* niente storage disponibile */
    }
    if (salvato) document.documentElement.setAttribute("data-theme", salvato);

    document.getElementById("theme").addEventListener("click", function () {
      var scuroOra =
        document.documentElement.getAttribute("data-theme") === "dark" ||
        (!document.documentElement.getAttribute("data-theme") &&
          window.matchMedia("(prefers-color-scheme: dark)").matches);
      var nuovo = scuroOra ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", nuovo);
      try {
        localStorage.setItem("tema", nuovo);
      } catch (err) {
        /* niente storage disponibile */
      }
    });
  })();

  // ------------------------------------------------------------ avvio

  function avvia(dati) {
    DATA = dati;
    var meta = document.getElementById("footer-meta");
    meta.textContent = "Dati generati da " + dati.source + " · " + dati.generated_at;
    window.addEventListener("hashchange", render);
    render();
  }

  if (window.__PLAN_DATA__) {
    avvia(window.__PLAN_DATA__);
  } else {
    fetch("data/plan.json", { cache: "no-cache" })
      .then(function (res) {
        if (!res.ok) throw new Error(res.status);
        return res.json();
      })
      .then(avvia)
      .catch(function () {
        app.innerHTML = "";
        app.appendChild(
          h("p", { class: "error", text: "Non riesco a caricare i dati del piano (data/plan.json)." })
        );
      });
  }
})();
