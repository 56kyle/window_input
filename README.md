# window_input
A set of tools for interacting with windows, primarily those that are non foreground and even minimized.

# Note
I'm currently in the process of rewriting this library in a much better manner under the new name [pywinput](https://github.com/56kyle/pywinput). The goal is to retain existing funcitonality, but in a much easier manner to work on that includes static typing and proper test driven devlopment. Additionally any keyboard and mouse related objects are going to try and mimick the syntax found in the mouse and keyboard libraries as much as possible.

### Installation
    pip install window_input

### Recommended Usage
    from window_input import Window, Key
    
### Examples

    from window_input import Window, Key
    import time
    
    if __name__ == "__main__":
        foreground = Window()
        foreground.key_down(Key.VK_MENU) # Alt key's name
        foreground.key_down(Key.VK_TAB)
        time.sleep(.2)
        foreground.key_up(Key.VK_MENU)
        foreground.key_up(Key.VK_TAB)
        
