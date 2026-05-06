# Iusprudentia Poenalis — Serveur MCP

Serveur MCP (Model Context Protocol) pour la recherche de jurisprudence du Tribunal fédéral suisse en droit pénal (CP/CPP), connecté directement à Claude.ai Pro.

---

## Contenu

| Fichier | Description |
|---|---|
| `server.py` | Serveur MCP Streamable HTTP avec gestion des sessions |
| `arrets.json` | Base de données : 4 729 arrêts TF 2021–2026 avec URLs |
| `requirements.txt` | Dépendances Python |
| `Procfile` | Configuration de démarrage pour Render |

---

## Outils exposés à Claude

| Outil | Description |
|---|---|
| `search_arrets` | Recherche par mots-clés, infraction, article de loi, année |
| `get_fulltext` | Texte intégral d'un arrêt depuis bger.ch |
| `get_references` | ATF et arrêts TF cités dans un arrêt |
| `get_arret_by_reference` | Texte d'un ATF ou arrêt TF cité en référence |

---

## Déploiement

### Étape 1 — GitHub

1. Créez un dépôt **privé** sur https://github.com
2. Uploadez les 4 fichiers via **Add file → Upload files**
3. Committez

### Étape 2 — Render.com

1. Créez un compte sur https://render.com
2. **New → Web Service** → connectez votre dépôt GitHub
3. Configurez :

| Champ | Valeur |
|---|---|
| Environment | `Python 3` |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn server:app --host 0.0.0.0 --port $PORT` |
| Instance Type | `Starter` ($7/mois) |

4. Cliquez **Create Web Service**
5. Attendez le message **Your service is live 🎉**

### Étape 3 — Connexion à Claude.ai

1. Ouvrez https://claude.ai (abonnement Pro requis)
2. **Settings → Connecteurs → Ajouter un connecteur personnalisé**
3. Entrez l'URL : `https://VOTRE-SERVICE.onrender.com/mcp`
4. Validez

### Étape 4 — Utilisation

Dans une nouvelle conversation Claude.ai :
1. Cliquez sur **+** dans la barre de saisie
2. Activez **Iusprudentia Poenalis**
3. Posez votre question juridique

**Exemples :**
- *"Quelles sont les conditions pour une expulsion selon l'art. 66a CP ?"*
- *"Recherche les arrêts récents sur la tentative de meurtre depuis 2024"*
- *"Quelle est la jurisprudence du TF sur la détention provisoire ?"*

---

## Mise à jour des données

Pour intégrer un nouveau fichier Excel :

1. Régénérez `arrets.json` avec le script de conversion (disponible sur demande)
2. Remplacez `arrets.json` sur GitHub
3. Render redéploie automatiquement en ~2 minutes

---

## Architecture technique

```
Claude.ai (Pro)
      ↕  MCP Streamable HTTP (POST /mcp)
      ↕  Gestion sessions (Mcp-Session-Id)
Render.com (Python/uvicorn)
      ↕
arrets.json (4 729 arrêts)  +  bger.ch (textes intégraux)
```

---

## Protocole

- Transport : **Streamable HTTP** (MCP spec 2025-11-25)
- Sessions : gérées via l'en-tête `Mcp-Session-Id`
- CORS : activé pour tous les domaines
