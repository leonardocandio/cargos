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
    
    def save_config(self, config: AppConfig) -> bool:
        """
        Save configuration to file.
        
        Args:
            config: AppConfig object to save
            
        Returns:
            True if successful, False otherwise
        """
        try:
            config_dict = {
                "destination_path": config.destination_path,
                "cargo_template_path": config.cargo_template_path,
                "autorizacion_template_path": config.autorizacion_template_path,
                "preview_rows_limit": config.preview_rows_limit
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
            
            config = AppConfig(
                destination_path=config_dict.get("destination_path", f"{DEFAULT_OUTPUT_DIR}/"),
                cargo_template_path=config_dict.get("cargo_template_path", f"{DEFAULT_TEMPLATES_DIR}/CARGO UNIFORMES.docx"),
                autorizacion_template_path=config_dict.get("autorizacion_template_path", f"{DEFAULT_TEMPLATES_DIR}/50% - AUTORIZACIÓN DESCUENTO DE UNIFORMES (02).docx"),
                preview_rows_limit=config_dict.get("preview_rows_limit", DEFAULT_PREVIEW_ROWS)
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
