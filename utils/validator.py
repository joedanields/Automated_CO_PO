"""
Validator Module
Validates consistency across multiple evaluation sheets
"""
from typing import Dict, List, Any, Optional, Union
from dataclasses import dataclass
from pathlib import Path
from io import BytesIO
from .data_parser import DataParser


@dataclass
class ValidationResult:
    """Result of validation check"""
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    
    def __str__(self):
        if self.is_valid:
            return "Validation Passed"
        return f"Validation Failed: {'; '.join(self.errors)}"


class Validator:
    """Validates evaluation sheets for consistency and correctness"""
    
    # Fields that must match across all evaluation sheets
    REQUIRED_MATCH_FIELDS = [
        'course_code',
        'course_name',
        'faculty_name',
        'regulation'
    ]
    
    # Fields that should match but can have warnings
    RECOMMENDED_MATCH_FIELDS = [
        'academic_year',
        'class_info'
    ]
    
    def __init__(self):
        """Initialize Validator"""
        self.parser = DataParser()
    
    def validate_file_exists(self, file_sources: List[Union[str, BytesIO]]) -> ValidationResult:
        """
        Validate that all files exist (for file paths) or are valid (for BytesIO)
        
        Args:
            file_sources: List of file paths or BytesIO objects to check
            
        Returns:
            ValidationResult
        """
        errors = []
        for source in file_sources:
            if isinstance(source, str) and not Path(source).exists():
                errors.append(f"File not found: {source}")
            elif isinstance(source, BytesIO) and source.closed:
                errors.append(f"BytesIO object is closed")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=[]
        )
    
    def validate_consistency(self, file_sources: List[Union[str, BytesIO]]) -> ValidationResult:
        """
        Validate that all evaluation sheets have matching metadata
        
        Args:
            file_sources: List of evaluation sheet file paths or BytesIO objects
            
        Returns:
            ValidationResult with validation status and any errors
        """
        if not file_sources:
            return ValidationResult(
                is_valid=False,
                errors=["No files provided for validation"],
                warnings=[]
            )
        
        errors = []
        warnings = []
        
        # Check files exist/valid
        file_check = self.validate_file_exists(file_sources)
        if not file_check.is_valid:
            return file_check
        
        # Extract metadata from all files
        all_metadata = []
        for idx, source in enumerate(file_sources):
            try:
                metadata = self.parser.extract_validation_fields(source)
                # Store identifier for error messages
                if isinstance(source, BytesIO):
                    metadata['file_identifier'] = getattr(source, 'name', f'File {idx+1}')
                else:
                    metadata['file_identifier'] = Path(source).name
                all_metadata.append(metadata)
            except Exception as e:
                source_name = getattr(source, 'name', f'File {idx+1}') if isinstance(source, BytesIO) else source
                errors.append(f"Error reading {source_name}: {str(e)}")
                return ValidationResult(is_valid=False, errors=errors, warnings=[])
        
        # Compare required fields
        reference = all_metadata[0]
        for i, metadata in enumerate(all_metadata[1:], start=1):
            for field in self.REQUIRED_MATCH_FIELDS:
                ref_value = reference.get(field, '').strip().upper()
                cur_value = metadata.get(field, '').strip().upper()
                
                if ref_value != cur_value:
                    errors.append(
                        f"Mismatch in '{field}': "
                        f"'{reference[field]}' (in {reference['file_identifier']}) vs "
                        f"'{metadata[field]}' (in {metadata['file_identifier']})"
                    )
        
        # Check recommended fields (warnings only)
        for i, metadata in enumerate(all_metadata[1:], start=1):
            for field in self.RECOMMENDED_MATCH_FIELDS:
                ref_value = reference.get(field, '').strip().upper()
                cur_value = metadata.get(field, '').strip().upper()
                
                if ref_value != cur_value:
                    warnings.append(
                        f"Difference in '{field}': "
                        f"'{reference[field]}' vs '{metadata[field]}'"
                    )
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
    
    def validate_regulation(self, file_sources: List[Union[str, BytesIO]], expected_regulation: str) -> ValidationResult:
        """
        Validate that all files match the expected regulation
        
        Args:
            file_sources: List of evaluation sheet file paths or BytesIO objects
            expected_regulation: Expected regulation (R17, R21, R24)
            
        Returns:
            ValidationResult
        """
        errors = []
        expected_norm = self.parser.normalize_regulation(expected_regulation)
        
        for idx, source in enumerate(file_sources):
            try:
                metadata = self.parser.extract_validation_fields(source)
                actual_reg = self.parser.normalize_regulation(metadata.get('regulation', ''))
                
                if actual_reg != expected_norm:
                    source_name = getattr(source, 'name', f'File {idx+1}') if isinstance(source, BytesIO) else Path(source).name
                    errors.append(
                        f"Regulation mismatch in {source_name}: "
                        f"expected {expected_norm}, found {actual_reg}"
                    )
            except Exception as e:
                source_name = getattr(source, 'name', f'File {idx+1}') if isinstance(source, BytesIO) else source
                errors.append(f"Error reading {source_name}: {str(e)}")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=[]
        )
    
    def validate_student_match(self, file_sources: List[Union[str, BytesIO]]) -> ValidationResult:
        """
        Validate that same students appear across all evaluation sheets
        
        Args:
            file_sources: List of evaluation sheet file paths or BytesIO objects
            
        Returns:
            ValidationResult with warnings for missing students
        """
        warnings = []
        
        all_students = []
        for idx, source in enumerate(file_sources):
            try:
                students = self.parser.extract_student_data(source)
                source_name = getattr(source, 'name', f'File {idx+1}') if isinstance(source, BytesIO) else Path(source).name
                all_students.append({
                    'file': source_name,
                    'reg_numbers': set(students.keys())
                })
            except Exception as e:
                source_name = getattr(source, 'name', f'File {idx+1}') if isinstance(source, BytesIO) else source
                warnings.append(f"Could not check students in {source_name}: {str(e)}")
        
        if len(all_students) < 2:
            return ValidationResult(is_valid=True, errors=[], warnings=warnings)
        
        # Find students missing from some sheets
        all_reg_numbers = set()
        for sheet in all_students:
            all_reg_numbers.update(sheet['reg_numbers'])
        
        for reg_no in all_reg_numbers:
            missing_from = []
            for sheet in all_students:
                if reg_no not in sheet['reg_numbers']:
                    missing_from.append(sheet['file'])
            
            if missing_from:
                warnings.append(
                    f"Student {reg_no} missing from: {', '.join(missing_from)}"
                )
        
        return ValidationResult(
            is_valid=True,  # Missing students is a warning, not error
            errors=[],
            warnings=warnings
        )
    
    def validate_marks_range(self, file_source: Union[str, BytesIO]) -> ValidationResult:
        """
        Validate that marks are within valid limits (0 to max)
        
        Args:
            file_source: Path to evaluation sheet or BytesIO object
            
        Returns:
            ValidationResult
        """
        errors = []
        warnings = []
        
        try:
            max_marks = self.parser.extract_max_marks(file_source)
            students = self.parser.extract_student_data(file_source)
            
            for reg_no, student in students.items():
                for co_num, mark in student['co_marks'].items():
                    max_mark = max_marks['co_max'].get(co_num, 0)
                    
                    if mark < 0:
                        errors.append(
                            f"Negative marks for {reg_no} in CO{co_num}: {mark}"
                        )
                    elif max_mark > 0 and mark > max_mark:
                        warnings.append(
                            f"Marks exceed max for {reg_no} in CO{co_num}: "
                            f"{mark} > {max_mark}"
                        )
        except Exception as e:
            errors.append(f"Error validating marks: {str(e)}")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
    
    def validate_all(self, file_sources: List[Union[str, BytesIO]], expected_regulation: str = None) -> ValidationResult:
        """
        Run all validation checks
        
        Args:
            file_sources: List of evaluation sheet file paths or BytesIO objects
            expected_regulation: Optional expected regulation
            
        Returns:
            Combined ValidationResult
        """
        all_errors = []
        all_warnings = []
        
        # File existence check
        result = self.validate_file_exists(file_sources)
        all_errors.extend(result.errors)
        if not result.is_valid:
            return ValidationResult(is_valid=False, errors=all_errors, warnings=[])
        
        # Consistency check
        result = self.validate_consistency(file_sources)
        all_errors.extend(result.errors)
        all_warnings.extend(result.warnings)
        
        # Regulation check
        if expected_regulation:
            result = self.validate_regulation(file_sources, expected_regulation)
            all_errors.extend(result.errors)
            all_warnings.extend(result.warnings)
        
        # Student match check
        result = self.validate_student_match(file_sources)
        all_warnings.extend(result.warnings)
        
        # Marks range check for each file
        for source in file_sources:
            result = self.validate_marks_range(source)
            all_errors.extend(result.errors)
            all_warnings.extend(result.warnings)
        
        return ValidationResult(
            is_valid=len(all_errors) == 0,
            errors=all_errors,
            warnings=all_warnings
        )


# Test the validator
if __name__ == '__main__':
    validator = Validator()
    
    # Test with sample files
    test_files = [
        'sample/input_R17/theory_eval/Dept_theory/C211_IA1_b1923_r17.xlsx',
        'sample/input_R17/theory_eval/Dept_theory/C211_ia2_B2023_R17.xlsx',
        'sample/input_R17/theory_eval/Dept_theory/C211_mod_B1923_R17.xlsx'
    ]
    
    print("=== Consistency Validation ===")
    result = validator.validate_consistency(test_files)
    print(f"Valid: {result.is_valid}")
    print(f"Errors: {result.errors}")
    print(f"Warnings: {result.warnings}")
    
    print("\n=== Regulation Validation ===")
    result = validator.validate_regulation(test_files, 'R17')
    print(f"Valid: {result.is_valid}")
    print(f"Errors: {result.errors}")
    
    print("\n=== Student Match Validation ===")
    result = validator.validate_student_match(test_files)
    print(f"Valid: {result.is_valid}")
    print(f"Warnings (first 5): {result.warnings[:5]}")
    
    print("\n=== Full Validation ===")
    result = validator.validate_all(test_files, 'R17')
    print(f"Valid: {result.is_valid}")
    print(f"Total Errors: {len(result.errors)}")
    print(f"Total Warnings: {len(result.warnings)}")
