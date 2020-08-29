# window_input
 A set of tools for interacting with windows, including those that are non foreground and even minimized.


### Installation
    pip install window_input

### Recommended Usage
    from window_input import Window, Key
    
### Examples

    from window_input import Window, Key
    import time
    
    if __name__ == "__main__":
        foreground = Window()
        foreground.key_down(Key.VK_F4)
        foreground.key_down(Key.VK_LALT)
        time.sleep(.2)
        foreground.key_up(Key.VK_F4)
        foreground.key_up(Key.VK_LALT)
