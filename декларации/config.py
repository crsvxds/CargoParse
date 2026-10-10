import os
import json

class Config:
    def __init__(self):
        self.config_file = os.path.join(os.path.dirname(__file__), "config.json")
        self.default_config = {
            "last_articles_path": "",
            "last_codes_dir": "",
            "last_output_dir": "",
            "last_invoice_prefix": "",
            "auto_open_results": True,
            "check_boxes": True,
            "remove_brackets": True,
            "dark_theme": False,
            "profiles": {
                "default": {
                    "articles_path": "",
                    "codes_dir": "",
                    "output_dir": ""
                }
            },
            "current_profile": "default"
        }
        self.config = self.load_config()

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    # Merge with defaults to handle new keys
                    config = self.default_config.copy()
                    config.update(loaded)
                    return config
            except Exception as e:
                print(f"Ошибка загрузки конфига: {e}")
                return self.default_config.copy()
        return self.default_config.copy()

    def save_config(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Ошибка сохранения конфига: {e}")
            return False

    def get(self, key, default=None):
        return self.config.get(key, default)

    def set(self, key, value):
        self.config[key] = value
        self.save_config()

    def get_profile(self, profile_name):
        return self.config.get("profiles", {}).get(profile_name, {})

    def save_profile(self, profile_name, profile_data):
        if "profiles" not in self.config:
            self.config["profiles"] = {}
        self.config["profiles"][profile_name] = profile_data
        self.save_config()

    def set_current_profile(self, profile_name):
        self.config["current_profile"] = profile_name
        self.save_config()

    def get_current_profile_data(self):
        current = self.config.get("current_profile", "default")
        return self.get_profile(current)
