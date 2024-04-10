# @staartwind.nl/lexical-paste-from-word
<sup><sub>Heavily inspired on ckeditor paste from word</sub></sup>

Pasting from programs like Word while maintaining style is a problem that many developers run against when putting some sort of rich text editor into practice.
We have developed a normalizer that converts the mess in a Word XML file into Lexical comprehensible plain HTML.

# Usage
```shell
yarn add @staartwind.nl/lexical-paste-from-word
```
Due to inconsistencies across Lexical versions and modifications to Lexical's handling of paste for rich text, we decided against creating a ready-to-use component.
Here's an example of how to change the original onPasteForRichText and listen for the PASTE_COMMAND.

## Current lexical function
Found on [Github](https://github.com/facebook/lexical/blob/main/packages/lexical-rich-text/src/index.ts#L433)
```typescript
function onPasteForRichText(
  event: CommandPayloadType<typeof PASTE_COMMAND>,
  editor: LexicalEditor,
): void {
  event.preventDefault();
  editor.update(
    () => {
      const selection = $getSelection();
      const clipboardData =
        objectKlassEquals(event, InputEvent) ||
        objectKlassEquals(event, KeyboardEvent)
          ? null
          : (event as ClipboardEvent).clipboardData;
      if (clipboardData != null && selection !== null) {
        $insertDataTransferForRichText(clipboardData, selection, editor);
      }
    },
    {
      tag: 'paste',
    },
  );
}
```

## Our implementation
```typescript
import {useLexicalComposerContext} from '@lexical/react/LexicalComposerContext'
import {useEffect} from 'react'
import {$getSelection, COMMAND_PRIORITY_CRITICAL, type CommandPayloadType, PASTE_COMMAND} from 'lexical'
import {objectKlassEquals} from '@lexical/utils'
import {$insertDataTransferForRichText} from '@lexical/clipboard'
import {MSWordNormalizer} from 'lexical-paste-from-word'

export default function ListenPastePlugin() {
    const [editor] = useLexicalComposerContext()

    const handlePaste = (event: CommandPayloadType<typeof PASTE_COMMAND>) => {
        event.preventDefault()
        editor.update(
            () => {
                const selection = $getSelection()
                const clipboardData =
                    objectKlassEquals(event, InputEvent) ||
                    objectKlassEquals(event, KeyboardEvent)
                        ? null
                        : (event as ClipboardEvent).clipboardData
                if (clipboardData != null && selection !== null) {
                    const data = clipboardData.getData('text/html')
                    const newData = new DataTransfer()
                    const wordNormalizer = new MSWordNormalizer()
                    newData.setData('text/html', data)
                    if (wordNormalizer.isActive(data)) {
                        newData.setData('text/html', wordNormalizer.normalize(data))
                    }
                    $insertDataTransferForRichText(newData, selection, editor)
                }
            },
            {
                tag: 'paste'
            }
        )
    }

    useEffect(() => {
        return editor.registerCommand(
            PASTE_COMMAND,
            (event) => {
                handlePaste(event)
                return true
            },
            COMMAND_PRIORITY_CRITICAL
        )
    }, [editor])

    return null
} 
```

And to use it
```typescript jsx
import ListenPastePlugin from '@/components/lexical/plugin/listenpasteplugin'

const LexicalEditor = () => {
    const initialConfig = {
        namespace: 'lexical-paste-from-word',
        nodes: []
    }
    return(
        <LexicalComposer initialConfig={initialConfig}>
            <TableContext>
                <div className={'relative rounded-xl border'}>
                    <RichTextPlugin
                        contentEditable={
                            <div
                                className={'editor-scroller min-h-[150px] border-0 flex relative outline-0 z-0 overflow-auto resize-y max-h-[500px]'}>
                                <div className={'flex-[auto] relative resize-y -z-[1]'} ref={onRef}>
                                    <ContentEditable/>
                                </div>
                            </div>}
                        placeholder={<></>}
                        ErrorBoundary={LexicalErrorBoundary}
                    />
                    <ListenPastePlugin/>
                </div>
            </TableContext>
        </LexicalComposer>
    )
}
```