### Asennus

```md robo-test```
```cd robo-test```
```code .```

open terminal

Luodaan uusi uv ympäristö. Jos Python versiota ei erikseen ilmoiteta tässä, niin se valitsee uusimman/oletusarvoisen niistä jotka ovat uv:lle ladattuna. Uusin versio jolla rpaframework tällä hetkellä toimii on 3.12.

```uv init –python 3.12```

Ladataan uv ympäristöön venv ympäristö ja lisätään siihen samalla pip asennustyökalu (eri kuin uv pip, joka on uv:n oma pip asennustapa). Pip asentuu tässä –seed parametrilla. Ilman tätä pip:n voisi asentaa myös erikseen käskyllä uv pip install pip

```uv venv –seed```

Aktivoidaan ympäristö käyttöön:

```.venv\Scripts\activate```

Asennetaan ohjelmistorobotiikkakirjasto:

```uv pip install rpaframework```
```uv pip install robocorp``` 


