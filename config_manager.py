"""
Configuration persistence manager for saving and loading user settings.
"""
import json
import logging
from pathlib import Path
from models import AppConfig
from constants import DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATES_DIR, DEFAULT_PREVIEW_ROWS


class ConfigManager:
    """Manages persistent configuration settings."""
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = Path(config_file)
        self.logger = logging.getLogger(__name__)
    
    def save_config(self, config: AppConfig, unified_config=None) -> bool:
        """
        Save configuration to consolidated config.json file.
        
        Args:
            config: AppConfig object to save
            unified_config: UnifiedConfig object to save (optional)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Load existing config to preserve unified_config if not provided
            existing_config = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    existing_config = json.load(f)
            
            # Prepare app settings
            app_settings = {
                "destination_path": str(Path(config.destination_path)),
                "cargo_template_path": str(Path(config.cargo_template_path)),
                "autorizacion_template_path": str(Path(config.autorizacion_template_path)),
                "preview_rows_limit": config.preview_rows_limit
            }
            
            # Prepare unified config (use provided or existing)
            if unified_config:
                unified_data = {
                    "occupations": [
                        {
                            "name": occ.name,
                            "display_name": occ.display_name,
                            "synonyms": occ.synonyms,
                            "prendas": [
                                {
                                    "prenda_type": prenda.prenda_type,
                                    "display_name": prenda.display_name,
                                    "has_sizes": prenda.has_sizes,
                                    "price_sml_other": prenda.price_sml_other,
                                    "price_xl_other": prenda.price_xl_other,
                                    "price_xxl_other": prenda.price_xxl_other,
                                    "price_sml_san_isidro": prenda.price_sml_san_isidro,
                                    "price_xl_san_isidro": prenda.price_xl_san_isidro,
                                    "price_xxl_san_isidro": prenda.price_xxl_san_isidro,
                                    "price_sml_tarapoto": prenda.price_sml_tarapoto,
                                    "price_xl_tarapoto": prenda.price_xl_tarapoto,
                                    "price_xxl_tarapoto": prenda.price_xxl_tarapoto,
                                }
                                for prenda in occ.prendas
                            ],
                            "is_active": occ.is_active
                        }
                        for occ in unified_config.occupations
                    ],
                    "default_occupation": unified_config.default_occupation,
                    "default_local_group": unified_config.default_local_group
                }
            else:
                # Use existing unified config if available
                unified_data = {
                    "occupations": existing_config.get("occupations", []),
                    "default_occupation": existing_config.get("default_occupation", "MOZO"),
                    "default_local_group": existing_config.get("default_local_group", "OTHER")
                }
            
            # Combine app settings and unified config
            config_dict = {
                "app_settings": app_settings,
                **unified_data
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Configuration saved to {self.config_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save configuration: {str(e)}")
            return False
    
    def load_config(self) -> AppConfig:
        """
        Load configuration from file.
        
        Returns:
            AppConfig object with loaded settings or defaults
        """
        try:
            if not self.config_file.exists():
                self.logger.info("No configuration file found, using defaults")
                return AppConfig()
            
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config_dict = json.load(f)
            
            # Load from consolidated config format
            app_settings = config_dict.get("app_settings", {})
            dest_path = app_settings.get("destination_path", f"{DEFAULT_OUTPUT_DIR}/")
            cargo_path = app_settings.get("cargo_template_path", f"{DEFAULT_TEMPLATES_DIR}/CARGO UNIFORMES.docx")
            autorizacion_path = app_settings.get("autorizacion_template_path", f"{DEFAULT_TEMPLATES_DIR}/50% - AUTORIZACIÓN DESCUENTO DE UNIFORMES (02).docx")
            preview_limit = app_settings.get("preview_rows_limit", DEFAULT_PREVIEW_ROWS)
            
            # Normalize paths for cross-platform compatibility
            if dest_path:
                # Convert Windows-style backslashes to forward slashes, then use pathlib
                dest_path = str(Path(dest_path.replace("\\", "/")))
            
            config = AppConfig(
                destination_path=dest_path,
                cargo_template_path=str(Path(cargo_path)),
                autorizacion_template_path=str(Path(autorizacion_path)),
                preview_rows_limit=preview_limit
            )
            
            self.logger.info(f"Configuration loaded from {self.config_file}")
            return config
            
        except Exception as e:
            self.logger.error(f"Failed to load configuration: {str(e)}")
            return AppConfig()
    
    def update_and_save(self, config: AppConfig, **updates) -> bool:
        """
        Update configuration values and save to file.
        
        Args:
            config: AppConfig object to update
            **updates: Key-value pairs to update
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Update config object
            for key, value in updates.items():
                if hasattr(config, key):
                    setattr(config, key, value)
            
            # Save updated config
            return self.save_config(config)
            
        except Exception as e:
            self.logger.error(f"Failed to update and save configuration: {str(e)}")
            return False
