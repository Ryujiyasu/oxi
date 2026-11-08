// SPDX-License-Identifier: MIT OR Apache-2.0

// Command gotenberg is the standard Gotenberg build plus the oxi module.
package main

import (
	gotenbergcmd "github.com/gotenberg/gotenberg/v9/cmd"
	_ "github.com/gotenberg/gotenberg/v9/pkg/standard"

	_ "github.com/Ryujiyasu/oxi/integrations/gotenberg"
)

func main() {
	gotenbergcmd.Run()
}
