# Listing images

How `plugin/images/cover.png`, `table-ja.png` and `timeline.png` were made. The decks are fictional, made with python-pptx. Each image shows the built plugin page (`dist/editor/index.html`) in Chrome, with a stand-in for `window.orca`. They show the plugin's own frame, not Orca's window, so no local path or account details are in them.

```sh
python3 make_decks.py en ws/q4-launch-plan.pptx
python3 make_decks.py ja ws/launch-plan-ja.pptx
python3 reorder.py ws/launch-plan-ja.pptx 3 srv/decks/budget-ja.pptx  # the table slide first
python3 reorder.py ws/q4-launch-plan.pptx 2 srv/decks/timeline.pptx   # the timeline first
cp ws/q4-launch-plan.pptx srv/decks/
ln -s ../../dist srv/plugin
python3 shoot.py srv q4-launch-plan.pptx cover.png 1600 1000 en-US
python3 shoot.py srv budget-ja.pptx table-ja.png 1600 1000 ja-JP
python3 shoot.py srv timeline.pptx timeline.png 1600 1000 en-US
```

`shoot.py` needs Playwright for Python and Google Chrome at `/usr/bin/google-chrome`. It lays the page out at 1100×688 and scales it to the requested width, so the slide fills the frame. The PNGs come out 1 px too tall; crop them to 1600×1000. Orca's limits: cover at least 800×500 and at most 1 MB; screenshots at least 1280×800 and at most 1.5 MB each; 8 MB for all images.
