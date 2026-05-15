import os
import sys

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def print_banner(title, subtitle=None):
    clear_screen()
    width = 70
    print("=" * width)
    print(title.center(width))
    if subtitle:
        print(subtitle.center(width))
    print("=" * width)

def get_input(prompt, default=None):
    val = input(f"{prompt} [{default}]: " if default else f"{prompt}: ").strip()
    return val if val else default
