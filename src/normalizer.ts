const MS_WORD_MATCHES = [
    /<meta\s*name="?generator"?\s*content="?microsoft\s*word\s*\d+"?\/?>/i,
    /xmlns:o="urn:schemas-microsoft-com/i
]

export class MSWordNormalizer {
    public isActive(htmlString: string): boolean {
        return MS_WORD_MATCHES.some(regex => regex.test(htmlString))
    }

    public normalize(htmlString: string): string {
        const {bodyString, stylesString} = parseHtml(htmlString)
        const doc = new DOMParser().parseFromString(bodyString, 'text/html')
        this.transformListItemLikeLElementsIntoLists(doc, stylesString)
        this.removeMSAttributes(doc)
        return doc.body.innerHTML
    }

    private findAllItemLikeElements(doc: Document): ListLikeElement[] {
        const items = doc.querySelectorAll('*')
        const res: ListLikeElement[] = []
        for (const item of items) {
            const tagName = item.tagName.toLowerCase()
            // Find all the possible list items
            if (!['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'div'].includes(tagName) || !(item instanceof HTMLElement)) {
                continue
            }

            const style = convertStyleToObject(item.getAttribute('style') || '')

            // Keep items that are part of a list.
            if ('mso-list' in style) {
                if (!item.parentElement) continue

                if (isList(item.parentElement!)) continue

                const itemData = getListItemData(style)

                res.push({
                    ...itemData,
                    element: item
                })
            }
        }

        return res
    }


    private transformListItemLikeLElementsIntoLists(doc: Document, stylesString: string) {
        const itemLikeElements = this.findAllItemLikeElements(doc)
        if (!itemLikeElements.length) {
            return
        }

        const encounteredLists: Record<string, number> = {}

        const stack: Array<ListLikeElement & {
            listElement: Element
            listItemElements: Array<Element>
        }> = []

        for (const itemLikeElement of itemLikeElements) {
            if (itemLikeElement.indent != null) {
                // Reset the current stack if the previous used list does not continue with itemLikeElement
                if ( !isListContinuation( itemLikeElement ) ) {
                    stack.length = 0;
                }
                // Combined list ID for addressing encounter lists counters.
                const originalListId = `${ itemLikeElement.id }:${ itemLikeElement.indent }`

                // Normalized list item indentation.
                const indent = Math.min( itemLikeElement.indent - 1, stack.length )

                // Trimming of the list stack on list ID change.
                if ( indent < stack.length && stack[ indent ]?.id !== itemLikeElement.id ) {
                    stack.length = indent
                }

                // Trimming of the list stack on lower indent list encountered.
                if ( indent < stack.length - 1 ) {
                    stack.length = indent + 1
                } else {
                    const listStyle = detectListStyle( itemLikeElement, stylesString )
                    // Create a new OL/UL if required (greater indent or different list type).
                    if ( indent > stack.length - 1 || stack[ indent ]?.listElement.tagName.toLowerCase() != listStyle.type ) {
                        if (
                            indent == 0 &&
                            listStyle.type == 'ol' &&
                            itemLikeElement.id !== undefined &&
                            encounteredLists[ originalListId ]
                        ) {
                            listStyle.startIndex = encounteredLists[ originalListId ] ?? null
                        }

                        const listElement = createNewEmptyList( listStyle, false )

                        // Insert the new OL/UL.
                        if ( stack.length == 0 ) {
                            const parent = itemLikeElement.element.parentNode!
                            const index = Array.from(parent.children).indexOf(itemLikeElement.element) + 1

                            if (index === parent.children.length) {
                                parent.appendChild(listElement)
                            } else {
                                parent.insertBefore(listElement, itemLikeElement.element)
                            }
                        } else {
                            const parentListItems = stack[ indent - 1 ]?.listItemElements

                            if (parentListItems) {
                                parentListItems[ parentListItems.length - 1 ]?.appendChild(listElement)
                            }
                        }

                        // Update the list stack for other items to reference.
                        stack[ indent ] = {
                            ...itemLikeElement,
                            listElement,
                            listItemElements: []
                        }

                        // Prepare list counter for start index.
                        if ( indent == 0 && itemLikeElement.id !== undefined ) {
                            encounteredLists[ originalListId ] = listStyle.startIndex || 1
                        }
                    }
                }

                // Use LI if it is already it or create a new LI element.
                const listItem = itemLikeElement.element.tagName.toLowerCase() == 'li' ? itemLikeElement.element : document.createElement( 'li' )

                // Append the LI to OL/UL.
                stack[indent]?.listElement.appendChild(listItem)
                stack[indent]?.listItemElements.push(listItem)

                // Increment list counter.
                if ( indent == 0 && itemLikeElement.id !== undefined ) {
                    encounteredLists[ originalListId ]++
                }

                // Append list block to LI.
                if ( itemLikeElement.element != listItem ) {
                    listItem.appendChild(itemLikeElement.element)
                }
                // Clean list block.
                removeMsoListIgnoreSpans( itemLikeElement.element )
            }
        }
    }

    private removeMSAttributes(doc: Document) {
        const items = doc.querySelectorAll('*')

        const elementsToUnwrap: Element[] = []

        for (const item of items) {
            // Eliminate all classes that include mso.
            for (const className of item.classList) {
                if ( /\bmso/gi.exec( className ) ) {
                    item.classList.remove(className)
                }
            }

            // Eliminate all useless mso styling
            if (item instanceof HTMLElement) {
                removeMSOStyles(item)
            }

            // Push word specific elements to an array
            if (item.tagName === 'W:SDT' || item.tagName === 'W:SDTPR' && isElementEmpty(item) || item.tagName === 'O:P' && isElementEmpty(item)) {
                elementsToUnwrap.push(item)
            }
        }

        // Eliminate word specific elements
        for (const item of elementsToUnwrap) {
            const itemParent = item.parentNode!
            itemParent.removeChild(item)
        }
    }
}

interface ListLikeElement extends ListItemData {
    element: Element
}


export function parseHtml(htmlString: string): {bodyString: string, stylesString: string} {
    const domParser = new DOMParser()

    // Remove Word specific "if comments" so content inside is not omitted by the parser.
    htmlString = htmlString.replace( /<!--\[if gte vml 1]>/g, '' )

    // Clean the <head> section of MS Windows specific tags. See https://github.com/ckeditor/ckeditor5/issues/15333.
    // The regular expression matches the <o:SmartTagType> tag with optional attributes (with or without values).
    htmlString = htmlString.replace( /<o:SmartTagType(?:\s+[^\s>=]+(?:="[^"]*")?)*\s*\/?>/gi, '' )

    const normalizedHtml = normalizeSpacing( cleanContentAfterBody( htmlString ) )

    // Parse htmlString as native Document object.
    const htmlDocument = domParser.parseFromString( normalizedHtml, 'text/html' )

    normalizeSpacerunSpans( htmlDocument )

    // Get `innerHTML` first as transforming to View modifies the source document.
    const bodyString = htmlDocument.body.innerHTML

    const stylesString = extractStyles(htmlDocument)

    return {
        bodyString,
        stylesString
    }
}

function extractStyles( htmlDocument: Document ): string {
    const styles = []
    const stylesString = []
    const styleTags = Array.from(htmlDocument.getElementsByTagName( 'style' ))

    for ( const style of styleTags ) {
        if ( style.sheet && style.sheet.cssRules && style.sheet.cssRules.length ) {
            styles.push( style.sheet )
            stylesString.push(style.innerHTML)
        }
    }

    return stylesString.join(' ')
}

export function normalizeSpacing( htmlString: string ): string {
    // Run normalizeSafariSpaceSpans() two times to cover nested spans.
    return normalizeSafariSpaceSpans( normalizeSafariSpaceSpans( htmlString ) )
        // Remove all \r\n from "spacerun spans" so the last replace line doesn't strip all whitespaces.
        .replace( /(<span\s+style=['"]mso-spacerun:yes['"]>[^\S\r\n]*?)[\r\n]+([^\S\r\n]*<\/span>)/g, '$1$2' )
        .replace( /<span\s+style=['"]mso-spacerun:yes['"]><\/span>/g, '' )
        .replace( /(<span\s+style=['"]letter-spacing:[^'"]+?['"]>)[\r\n]+(<\/span>)/g, '$1 $2' )
        .replace( / <\//g, '\u00A0</' )
        .replace( / <o:p><\/o:p>/g, '\u00A0<o:p></o:p>' )
        // Remove <o:p> block filler from empty paragraph. Safari uses \u00A0 instead of &nbsp;.
        .replace( /<o:p>(&nbsp;|\u00A0)<\/o:p>/g, '' )
        // Remove all whitespaces when they contain any \r or \n.
        .replace( />([^\S\r\n]*[\r\n]\s*)</g, '><' )
}

function normalizeSafariSpaceSpans( htmlString: string ) {
    return htmlString.replace( /<span(?: class="Apple-converted-space"|)>(\s+)<\/span>/g, ( fullMatch, spaces ) => {
        return spaces.length === 1 ? ' ' : Array( spaces.length + 1 ).join( '\u00A0 ' ).substring( 0, spaces.length )
    })
}

export function normalizeSpacerunSpans( htmlDocument: Document ): void {
    htmlDocument.querySelectorAll( 'span[style*=spacerun]' ).forEach( el => {
        const htmlElement = el as HTMLElement
        const innerTextLength = htmlElement.innerText.length || 0

        htmlElement.innerText = Array( innerTextLength + 1 ).join( '\u00A0 ' ).substring( 0, innerTextLength )
    })
}

function cleanContentAfterBody( htmlString: string ) {
    const bodyCloseTag = '</body>'
    const htmlCloseTag = '</html>'

    const bodyCloseIndex = htmlString.indexOf( bodyCloseTag )

    if ( bodyCloseIndex < 0 ) {
        return htmlString
    }

    const htmlCloseIndex = htmlString.indexOf( htmlCloseTag, bodyCloseIndex + bodyCloseTag.length )

    return htmlString.substring( 0, bodyCloseIndex + bodyCloseTag.length ) +
        ( htmlCloseIndex >= 0 ? htmlString.substring( htmlCloseIndex ) : '' )
}

function convertStyleToObject(styleString: string): Record<string, string> {
    const styleObject: Record<string, string> = {}

    if (!styleString) {
        return styleObject
    }

    // Split the style string into individual declarations
    const declarations = styleString.split(';').map(declaration => declaration.trim())

    // Iterate over each declaration and split it into property and value
    declarations.forEach(declaration => {
        const [property, value] = declaration.split(':').map(part => part.trim())
        if (property && value) {
            // Add the property and value to the style object
            styleObject[property] = value
        }
    })

    return styleObject
}

function isElementEmpty(element: Element): boolean {
    return !element.textContent && !element.innerHTML.trim()
}

function removeMSOStyles(element: HTMLElement): void {
    // Get the current style string of the paragraph node
    const currentStyle = element.getAttribute('style') || ''

    // Parse the style string to extract individual style properties
    const styleProperties = currentStyle.split(';').map(property => property.trim())

    // Remove the 'mso-bidi-font-family' property from the list of style properties
    const filteredProperties = styleProperties.filter(property => {
        const [key] = property.split(':')
        return !(/\bmso/gi.exec(key?.toLowerCase() ?? ''))
    })

    // Reconstruct the style string without the removed property
    const updatedStyle = filteredProperties.join(';')

    // Set the updated style string back to the paragraph node
    element.setAttribute('style', updatedStyle)
}

interface ListItemData {
    id?: string
    order?: string
    indent?: number
}

function getListItemData(styles: {[key: string]: string}): ListItemData {
    const listStyle = styles['mso-list']

    if ( listStyle === undefined ) {
        return {}
    }

    const idMatch = listStyle.match( /(^|\s{1,100})l(\d+)/i )
    const orderMatch = listStyle.match( /\s{0,100}lfo(\d+)/i )
    const indentMatch = listStyle.match( /\s{0,100}level(\d+)/i )

    if ( idMatch && orderMatch && indentMatch ) {
        return {
            id: idMatch[2],
            order: orderMatch[ 1 ],
            indent: parseInt( indentMatch[ 1 ] ?? '' )
        }
    }

    return {
        indent: 1
    }
}

function detectListStyle( listLikeItem: ListLikeElement, stylesString: string ) {
    const listStyleRegexp = new RegExp( `@list l${ listLikeItem.id }:level${ listLikeItem.indent }\\s*({[^}]*)`, 'gi' )
    const listStyleTypeRegex = /mso-level-number-format:([^;]{0,100});/gi
    const listStartIndexRegex = /mso-level-start-at:\s{0,100}([0-9]{0,10})\s{0,100};/gi
    const legalStyleListRegex = new RegExp( `@list\\s+l${ listLikeItem.id }:level\\d\\s*{[^{]*mso-level-text:"%\\d\\\\.`, 'gi' )
    const multiLevelNumberFormatTypeRegex = new RegExp( `@list l${ listLikeItem.id }:level\\d\\s*{[^{]*mso-level-number-format:`, 'gi' )

    const legalStyleListMatch = legalStyleListRegex.exec( stylesString )
    const multiLevelNumberFormatMatch = multiLevelNumberFormatTypeRegex.exec( stylesString )

    // Multi level lists in Word have mso-level-number-format attribute except legal lists,
    // so we used that. If list has legal list match and doesn't has mso-level-number-format
    // then this is legal-list.
    const islegalStyleList = legalStyleListMatch && !multiLevelNumberFormatMatch

    const listStyleMatch = listStyleRegexp.exec( stylesString )

    let listStyleType = 'decimal' // Decimal is default one.
    let type = 'ol' // <ol> is default list.
    let startIndex = null

    if ( listStyleMatch && listStyleMatch[ 1 ] ) {
        const listStyleTypeMatch = listStyleTypeRegex.exec( listStyleMatch[ 1 ] )

        if ( listStyleTypeMatch && listStyleTypeMatch[ 1 ] ) {
            listStyleType = listStyleTypeMatch[ 1 ].trim()
            type = listStyleType !== 'bullet' && listStyleType !== 'image' ? 'ol' : 'ul'
        }

        // Styles for the numbered lists are always defined in the Word CSS stylesheet.
        // Unordered lists MAY contain a value for the Word CSS definition `mso-level-text` but sometimes
        // this tag is missing. And because of that, we cannot depend on that. We need to predict the list style value
        // based on the list style marker element.
        if ( listStyleType === 'bullet' ) {
            const bulletedStyle = findBulletedListStyle( listLikeItem.element )

            if ( bulletedStyle ) {
                listStyleType = bulletedStyle
            }
        } else {
            const listStartIndexMatch = listStartIndexRegex.exec( listStyleMatch[ 1 ] )

            if ( listStartIndexMatch && listStartIndexMatch[ 1 ] ) {
                startIndex = parseInt( listStartIndexMatch[ 1 ] )
            }
        }

        if ( islegalStyleList ) {
            type = 'ol'
        }
    }

    return {
        type,
        startIndex,
        style: mapListStyleDefinition( listStyleType ),
        isLegalStyleList: islegalStyleList
    }
}

function findBulletedListStyle( element: Element ) {
    // https://github.com/ckeditor/ckeditor5/issues/15964
    if ( element.tagName.toLowerCase() == 'li' && element.parentElement!.tagName.toLowerCase() == 'ul' && element.parentElement!.hasAttribute( 'type' ) ) {
        return element.parentElement!.getAttribute( 'type' )
    }

    const listMarkerElement = findListMarkerNode( element )

    if ( !listMarkerElement ) {
        return null
    }


    const listMarker = listMarkerElement.textContent

    if ( listMarker === 'o' ) {
        return 'circle'
    } else if ( (listMarker ?? '').charCodeAt(0) === 183 ) {
        return 'disc'
    }
    // Word returns '§' instead of '■' for the square list style.
    else if ( listMarker === '§' ) {
        return 'square'
    }

    return null
}

function findListMarkerNode( element: Element ): Element | null {
    // If the first child is a text node, it is the data for the element.
    // The list-style marker is not present here.
    if(Array.from(element.children)[0]?.tagName.toLowerCase() === '#text') {
        return null
    }

    for ( const childNode of Array.from(element.children) ) {
        // The list-style marker will be inside the `<span>` element. Let's ignore all non-span elements.
        // It may happen that the `<a>` element is added as the first child. Most probably, it's an anchor element.
        if ( !['element', 'span'].includes(childNode.tagName.toLowerCase())) {
            continue
        }

        return childNode
    }

    return null
}

function mapListStyleDefinition( value: string ) {
    if ( value.startsWith( 'arabic-leading-zero' ) ) {
        return 'decimal-leading-zero'
    }

    switch ( value ) {
        case 'alpha-upper':
            return 'upper-alpha'
        case 'alpha-lower':
            return 'lower-alpha'
        case 'roman-upper':
            return 'upper-roman'
        case 'roman-lower':
            return 'lower-roman'
        case 'circle':
        case 'disc':
        case 'square':
            return value
        default:
            return null
    }
}

function createNewEmptyList(
    listStyle: ReturnType<typeof detectListStyle>,
    hasMultiLevelListPlugin: boolean
) {
    const list = document.createElement(listStyle.type)

    // We do not support modifying the marker for a particular list item.
    // Set the value for the `list-style-type` property directly to the list container.
    if ( listStyle.style ) {
        list.style.setProperty('list-style-type', listStyle.style)
    }

    if ( listStyle.startIndex && listStyle.startIndex > 1 ) {
        list.setAttribute('start', listStyle.startIndex.toString())
    }

    if ( listStyle.isLegalStyleList && hasMultiLevelListPlugin ) {
        list.classList.add('legal-list')
    }

    return list
}

function removeMsoListIgnoreSpans(parentElement: Element) {
    // Get all <span> elements inside the parentElement
    const spans = parentElement.querySelectorAll('span')

    // Iterate over each <span> element
    spans.forEach(span => {
        // Get the value of the 'style' attribute
        const style = span.getAttribute('style')
        // Check if the 'style' attribute contains 'mso-list: Ignore'
        if (style && style.includes('mso-list:Ignore')) {
            // Remove the <span> element
            span.parentNode?.removeChild(span)
        }
    })
}

function isListContinuation( currentItem: ListLikeElement ) {
    const previousSibling = currentItem.element.previousSibling;

    if ( !previousSibling ) {
        // If it's a li inside ul or ol like in here: https://github.com/ckeditor/ckeditor5/issues/15964.
        return isList( currentItem.element.parentElement! )
    }

    // Even with the same id the list does not have to be continuous (#43).
    // @ts-expect-error
    return isList( previousSibling )
}

const isList = (element: HTMLElement) => ['ul', 'ol'].includes(element.tagName.toLowerCase())
