"""
Template Mapper Module
Maps regulation, category, and department type to correct template files
"""
import os
from pathlib import Path


class TemplateMapper:
    """Maps regulation + category + dept_type to correct template file"""
    
    # Template mapping configuration
    TEMPLATE_MAP = {
        'R17': {
            'theory': {
                'dept': 'Dept THEORY template_ R17 V3 AtSheet.xlsx',
                's&h': 'S&H THEORY template _R17 V3 AtSheet.xlsx'
            },
            'analytical': {
                'dept': 'Dept THEORY Analytical template_R17 V3 AtSheet.xlsx',
                's&h': 'S&H THEORY template Analytical_R17 V3 AtSheet.xlsx'
            },
            'lab': {
                'default': 'LAB template_R17 V3 AtSheet.xlsx'
            },
            'project': {
                'default': 'Project template_R17 V3 AtSheet.xlsx'
            }
        },
        'R21': {
            'theory': {
                'dept': 'Dept THEORY  template_R21 V1 AtSheet.xlsx',
                's&h': 'Dept THEORY  template_R21 V1 AtSheet.xlsx'  # Using same for now
            },
            'analytical': {
                'dept': 'Dept THEORY Analytical template_R21 V1 AtSheet.xlsx',
                's&h': 'Dept THEORY Analytical template_R21 V1 AtSheet.xlsx'
            },
            'integrated': {
                'dept': 'Dept Integrated template_R21 V1 AtSheet.xlsx',
                's&h': 'Dept Integrated template_R21 V1 AtSheet.xlsx'
            },
            'lab': {
                'default': 'LAB template_R21 V1AtSheet.xlsx'
            },
            'project': {
                'default': 'Project template_R21 V1 AtSheet.xlsx'
            }
        },
        'R24': {
            'theory': {
                'dept': 'Dept THEORY  template_R21 V1 AtSheet.xlsx',
                's&h': 'Dept THEORY  template_R21 V1 AtSheet.xlsx'
            },
            'analytical': {
                'dept': 'Dept THEORY Analytical template_R21 V1 AtSheet.xlsx',
                's&h': 'Dept THEORY Analytical template_R21 V1 AtSheet.xlsx'
            },
            'integrated': {
                'dept': 'Dept Integrated template_R21 V1 AtSheet.xlsx',
                's&h': 'Dept Integrated template_R21 V1 AtSheet.xlsx'
            },
            'lab': {
                'default': 'LAB template_R21 V1AtSheet.xlsx'
            },
            'project': {
                'default': 'Project template_R21 V1 AtSheet.xlsx'
            }
        }
    }
    
    # Required input files for each category
    REQUIRED_INPUTS = {
        'R17': {
            'theory': ['IA1', 'IA2', 'Model'],
            'analytical': ['IA1', 'IA2', 'Model'],
            'lab': ['Lab'],
            'project': ['Review1', 'Review2', 'Review3']
        },
        'R21': {
            'theory': ['IA1', 'IA2', 'Integrated'],
            'analytical': ['IA1', 'IA2', 'Integrated'],
            'integrated': ['IA1', 'IA2', 'Integrated'],
            'lab': ['Lab'],
            'project': ['Review1', 'Review2', 'Review3']
        },
        'R24': {
            'theory': ['IA1', 'IA2', 'Integrated'],
            'analytical': ['IA1', 'IA2', 'Integrated'],
            'integrated': ['IA1', 'IA2', 'Integrated'],
            'lab': ['Lab'],
            'project': ['Review1', 'Review2', 'Review3']
        }
    }
    
    def __init__(self, base_path: str = None):
        """
        Initialize TemplateMapper
        
        Args:
            base_path: Base path to the project directory
        """
        if base_path is None:
            self.base_path = Path(__file__).parent.parent
        else:
            self.base_path = Path(base_path)
        
        self.template_dir = self.base_path / 'Attainment_Template'
    
    def get_regulation_folder(self, regulation: str) -> str:
        """
        Get folder name for regulation
        
        Args:
            regulation: Regulation code (R17, R21, R24)
            
        Returns:
            Folder name (Reg_17, Reg_21, Reg_24)
        """
        reg_map = {
            'R17': 'Reg_17',
            'R21': 'Reg_21',
            'R24': 'Reg_24'
        }
        return reg_map.get(regulation.upper(), 'Reg_17')
    
    def get_template_path(self, regulation: str, category: str, dept_type: str = 'default') -> Path:
        """
        Get path to the correct template file
        
        Args:
            regulation: Regulation code (R17, R21, R24)
            category: Course category (theory, analytical, lab, project)
            dept_type: Department type (dept, s&h, default)
            
        Returns:
            Path to template file
            
        Raises:
            ValueError: If template not found for given parameters
        """
        regulation = regulation.upper()
        category = category.lower()
        dept_type = dept_type.lower()
        
        if regulation not in self.TEMPLATE_MAP:
            raise ValueError(f"Unknown regulation: {regulation}. Valid options: {list(self.TEMPLATE_MAP.keys())}")
        
        if category not in self.TEMPLATE_MAP[regulation]:
            raise ValueError(f"Unknown category: {category} for {regulation}. Valid options: {list(self.TEMPLATE_MAP[regulation].keys())}")
        
        category_templates = self.TEMPLATE_MAP[regulation][category]
        
        # Get template filename
        if dept_type in category_templates:
            template_filename = category_templates[dept_type]
        elif 'default' in category_templates:
            template_filename = category_templates['default']
        else:
            raise ValueError(f"Unknown department type: {dept_type} for {regulation}/{category}")
        
        # Construct full path
        reg_folder = self.get_regulation_folder(regulation)
        template_path = self.template_dir / reg_folder / template_filename
        
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        return template_path
    
    def get_required_inputs(self, regulation: str, category: str) -> list:
        """
        Get list of required input files for given regulation and category
        
        Args:
            regulation: Regulation code (R17, R21, R24)
            category: Course category (theory, analytical, lab, project)
            
        Returns:
            List of required input types (e.g., ['IA1', 'IA2', 'Model'])
        """
        regulation = regulation.upper()
        category = category.lower()
        
        if regulation not in self.REQUIRED_INPUTS:
            raise ValueError(f"Unknown regulation: {regulation}")
        
        if category not in self.REQUIRED_INPUTS[regulation]:
            raise ValueError(f"Unknown category: {category} for {regulation}")
        
        return self.REQUIRED_INPUTS[regulation][category]
    
    def get_available_regulations(self) -> list:
        """Get list of available regulations"""
        return list(self.TEMPLATE_MAP.keys())
    
    def get_available_categories(self, regulation: str) -> list:
        """Get list of available categories for a regulation"""
        regulation = regulation.upper()
        if regulation not in self.TEMPLATE_MAP:
            return []
        return list(self.TEMPLATE_MAP[regulation].keys())
    
    def get_available_dept_types(self, regulation: str, category: str) -> list:
        """Get list of available department types for a regulation and category"""
        regulation = regulation.upper()
        category = category.lower()
        
        if regulation not in self.TEMPLATE_MAP:
            return []
        if category not in self.TEMPLATE_MAP[regulation]:
            return []
        
        dept_types = list(self.TEMPLATE_MAP[regulation][category].keys())
        # Remove 'default' from the list if present and return meaningful options
        if 'default' in dept_types and len(dept_types) == 1:
            return ['default']
        return [d for d in dept_types if d != 'default']


# Test the mapper
if __name__ == '__main__':
    mapper = TemplateMapper()
    
    print("Available Regulations:", mapper.get_available_regulations())
    
    for reg in mapper.get_available_regulations():
        print(f"\n{reg} Categories:", mapper.get_available_categories(reg))
        for cat in mapper.get_available_categories(reg):
            print(f"  {cat} - Dept Types:", mapper.get_available_dept_types(reg, cat))
            print(f"  {cat} - Required Inputs:", mapper.get_required_inputs(reg, cat))
