#!/usr/bin/env python3
"""
Microsoft Teams Chat Export Script (Legacy Wrapper)

This is a legacy wrapper that maintains backward compatibility.
The main functionality has been moved to teams_chat_export.py with improved structure.

For new usage, consider using: python teams_chat_export.py

Author: Teams Chat Export Script
Version: v0.1.1
"""

import sys
import os

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from teams_chat_export import main
    
    print("⚠️  Note: You are using the legacy script wrapper.")
    print("   Consider using 'python teams_chat_export.py' for the new structured version.")
    print("   Functionality remains the same.\n")
    
    # Run the main export function
    main()
    
except ImportError as e:
    print(f"❌ Error importing required modules: {e}")
    print("   Please ensure all required files are present:")
    print("   - teams_chat_export.py")
    print("   - teams_utils.py") 
    print("   - config.py")
    print("   - htmlTemplate.py")
    sys.exit(1)
except Exception as e:
    print(f"❌ Unexpected error: {e}")
    sys.exit(1)
