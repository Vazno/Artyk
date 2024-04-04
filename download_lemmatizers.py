import os
import pathlib

import spacy
import spacy.cli.download


def download_model(model_name):
    spacy.cli.download(model_name)

    nlp = spacy.load(model_name)
    if not pathlib.Path(os.path.join("models", model_name)).exists():
        os.mkdir(os.path.join("models", model_name))

    nlp.to_disk(os.path.join("models", model_name))

models = {
    "croatian": "hr_core_news_sm",
    "danish": "da_core_news_sm",
    "dutch": "nl_core_news_sm",
    "english": "en_core_web_sm", 
    "finnish": "fi_core_news_sm",
    "german": "de_core_news_sm",
    "bokmal": "nb_core_news_sm",
    "russian": "ru_core_news_sm",
    "ukranian": "uk_core_news_sm",
}
if __name__ == "__main__":
    for model, name in models:
        download_model(model)